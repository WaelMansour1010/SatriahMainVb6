VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmSearchinvestment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   7380
   ClientLeft      =   4665
   ClientTop       =   4335
   ClientWidth     =   13785
   Icon            =   "FrmSearchinvestment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   13785
   Begin VB.Frame Frame22 
      BackColor       =   &H00E2E9E9&
      Height          =   1695
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   253
      Top             =   4410
      Visible         =   0   'False
      Width           =   13425
      Begin VB.TextBox txtContainerNo 
         BackColor       =   &H0000FFFF&
         Height          =   345
         Left            =   120
         TabIndex        =   286
         Top             =   1320
         Width           =   1725
      End
      Begin VB.TextBox Text29 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   285
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TxtLeaderName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   274
         Top             =   1350
         Width           =   1395
      End
      Begin VB.TextBox TxtManualNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6240
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   266
         Top             =   240
         Width           =   915
      End
      Begin MSDataListLib.DataCombo DcbBranch22 
         Bindings        =   "FrmSearchinvestment.frx":6852
         Height          =   315
         Left            =   8400
         TabIndex        =   254
         Top             =   240
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   8400
         TabIndex        =   256
         Top             =   600
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbShip 
         Height          =   315
         Left            =   8400
         TabIndex        =   258
         Top             =   960
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCityFromId 
         Height          =   315
         Left            =   10590
         TabIndex        =   260
         Top             =   1320
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCityToId 
         Height          =   315
         Left            =   8400
         TabIndex        =   262
         Top             =   1320
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo VehicleType 
         Height          =   315
         Left            =   4200
         TabIndex        =   264
         Top             =   240
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCCar 
         Height          =   315
         Left            =   4200
         TabIndex        =   268
         Top             =   600
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCar2 
         Height          =   315
         Left            =   120
         TabIndex        =   270
         Top             =   600
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCEmp 
         Height          =   315
         Left            =   5580
         TabIndex        =   272
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbSupplem 
         Height          =   315
         Left            =   4200
         TabIndex        =   276
         Top             =   960
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbSupplem2 
         Height          =   315
         Left            =   120
         TabIndex        =   278
         Top             =   960
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton ChCarType 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   281
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "„„·Êﬂ… ··‘—ﬂ…"
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton ChCarType 
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   282
         Top             =   240
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "„„·Êﬂ… ··€Ì—"
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton ChCarType 
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   284
         Top             =   240
         Width           =   615
         _Version        =   786432
         _ExtentX        =   1085
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «—«„ﬂÊ"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   1830
         TabIndex        =   287
         Top             =   1410
         Width           =   1050
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·„—ﬂ»…"
         Height          =   285
         Index           =   81
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   283
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„·Õﬁ"
         Height          =   285
         Index           =   75
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   279
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„·Õﬁ"
         Height          =   285
         Index           =   74
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   277
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "”«∆ﬁ Œ«—ÃÌ"
         Height          =   285
         Index           =   73
         Left            =   4470
         RightToLeft     =   -1  'True
         TabIndex        =   275
         Top             =   1350
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Õœœ «·”«∆ﬁ"
         Height          =   285
         Index           =   72
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   273
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„„·ÊﬂÂ ··€Ì—"
         Height          =   285
         Index           =   71
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   271
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„„·Êﬂ… ··‘—ﬂ…"
         Height          =   285
         Index           =   70
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   269
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
         Height          =   285
         Index           =   82
         Left            =   7140
         RightToLeft     =   -1  'True
         TabIndex        =   267
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ—«“ «·„—ﬂ»…"
         Height          =   285
         Index           =   69
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   265
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   68
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   263
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·—Õ·… „‰ "
         Height          =   285
         Index           =   67
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   261
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·”›Ì‰…"
         Height          =   285
         Index           =   66
         Left            =   12060
         RightToLeft     =   -1  'True
         TabIndex        =   259
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì·"
         Height          =   285
         Index           =   59
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   257
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   31
         Left            =   12390
         RightToLeft     =   -1  'True
         TabIndex        =   255
         Top             =   240
         Width           =   1050
      End
   End
   Begin VB.Frame Frame18 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   220
      Top             =   4440
      Width           =   13425
      Begin VB.Frame Frame19 
         BackColor       =   &H00E2E9E9&
         Height          =   1575
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   232
         Top             =   0
         Visible         =   0   'False
         Width           =   13425
         Begin VB.TextBox Summary 
            Alignment       =   1  'Right Justify
            Height          =   570
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   246
            Top             =   720
            Width           =   2985
         End
         Begin VB.Frame Frame21 
            BackColor       =   &H00E2E9E9&
            Caption         =   "› —… «·Œ—ÊÃ "
            Height          =   615
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   241
            Top             =   840
            Width           =   5415
            Begin MSComCtl2.DTPicker FrmExitDate 
               Height          =   330
               Left            =   2760
               TabIndex        =   242
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
               Format          =   103546883
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker ToExitDate 
               Height          =   330
               Left            =   120
               TabIndex        =   243
               Top             =   270
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
               Format          =   103546883
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   64
               Left            =   4890
               RightToLeft     =   -1  'True
               TabIndex        =   245
               Top             =   360
               Width           =   300
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   63
               Left            =   2310
               RightToLeft     =   -1  'True
               TabIndex        =   244
               Top             =   360
               Width           =   360
            End
         End
         Begin VB.Frame Frame20 
            BackColor       =   &H00E2E9E9&
            Caption         =   "› —… «·œŒÊ·"
            Height          =   615
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   236
            Top             =   120
            Width           =   5415
            Begin MSComCtl2.DTPicker FrmEnterDate 
               Height          =   330
               Left            =   2760
               TabIndex        =   237
               Top             =   270
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
               Format          =   103546883
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker ToEnterDate 
               Height          =   330
               Left            =   120
               TabIndex        =   238
               Top             =   270
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/MM/yyyy,hh:mm:tt"
               Format          =   103546883
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   62
               Left            =   2190
               RightToLeft     =   -1  'True
               TabIndex        =   240
               Top             =   240
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   61
               Left            =   4770
               RightToLeft     =   -1  'True
               TabIndex        =   239
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.TextBox TxtNoImpExp 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   235
            Top             =   600
            Width           =   2625
         End
         Begin VB.TextBox Txtbarcode 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   234
            Top             =   960
            Width           =   2625
         End
         Begin VB.ComboBox DcbImportExport 
            Height          =   315
            ItemData        =   "FrmSearchinvestment.frx":6867
            Left            =   5640
            List            =   "FrmSearchinvestment.frx":6869
            RightToLeft     =   -1  'True
            TabIndex        =   233
            Top             =   240
            Width           =   2985
         End
         Begin MSDataListLib.DataCombo ArchSearchBranchDC 
            Bindings        =   "FrmSearchinvestment.frx":686B
            Height          =   315
            Left            =   9840
            TabIndex        =   247
            Top             =   240
            Width           =   2625
            _ExtentX        =   4630
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·Œ÷ «·”‰œ"
            Height          =   300
            Index           =   65
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   252
            Top             =   840
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   300
            Index           =   60
            Left            =   12615
            TabIndex        =   251
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»«—ﬂÊœ"
            Height          =   285
            Index           =   56
            Left            =   12525
            TabIndex        =   250
            Top             =   960
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰Ê⁄"
            Height          =   300
            Index           =   55
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   249
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   30
            Left            =   12390
            RightToLeft     =   -1  'True
            TabIndex        =   248
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.TextBox TxtTelephone 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   228
         Top             =   480
         Width           =   4665
      End
      Begin VB.TextBox TxtNameP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   225
         Top             =   480
         Width           =   4680
      End
      Begin VB.TextBox projTxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   10950
         RightToLeft     =   -1  'True
         TabIndex        =   221
         Top             =   840
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo DcbKBranch 
         Bindings        =   "FrmSearchinvestment.frx":6880
         Height          =   315
         Left            =   840
         TabIndex        =   222
         Top             =   840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbProject 
         Bindings        =   "FrmSearchinvestment.frx":6895
         Height          =   315
         Left            =   7320
         TabIndex        =   223
         Top             =   840
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Â« ›"
         Height          =   285
         Index           =   54
         Left            =   5895
         TabIndex        =   229
         Top             =   480
         Width           =   750
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘—Ê⁄"
         Height          =   285
         Index           =   58
         Left            =   12360
         TabIndex        =   227
         Top             =   840
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·Õ«Ã“"
         Height          =   285
         Index           =   57
         Left            =   12225
         TabIndex        =   226
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   29
         Left            =   5790
         RightToLeft     =   -1  'True
         TabIndex        =   224
         Top             =   840
         Width           =   1050
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   206
      Top             =   4440
      Width           =   13425
      Begin VB.Frame Frame17 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   213
         Top             =   120
         Width           =   5895
         Begin MSComCtl2.DTPicker FromLiqDate 
            Height          =   330
            Left            =   2280
            TabIndex        =   214
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   103546883
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker ToLiqDate 
            Height          =   330
            Left            =   120
            TabIndex        =   215
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   103546883
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï "
            Height          =   315
            Index           =   53
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ "
            Height          =   315
            Index           =   52
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ’›Ì…"
            Height          =   195
            Index           =   51
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   240
            Width           =   1425
         End
      End
      Begin VB.TextBox TxtDcbEmploSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11070
         RightToLeft     =   -1  'True
         TabIndex        =   211
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox TxtCodeInvest 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11070
         RightToLeft     =   -1  'True
         TabIndex        =   209
         Top             =   600
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo DcbInvest 
         Bindings        =   "FrmSearchinvestment.frx":68AA
         Height          =   315
         Left            =   6120
         TabIndex        =   210
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbEmployee 
         Bindings        =   "FrmSearchinvestment.frx":68BF
         Height          =   315
         Left            =   6120
         TabIndex        =   212
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„”«Â„…"
         Height          =   285
         Index           =   34
         Left            =   11790
         RightToLeft     =   -1  'True
         TabIndex        =   208
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ«∆„ »«· ’›Ì… "
         Height          =   285
         Index           =   33
         Left            =   11940
         RightToLeft     =   -1  'True
         TabIndex        =   207
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   184
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox Text28 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10320
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   202
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox Text27 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   201
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox Text25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   198
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox Text22 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   197
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox Text26 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   10680
         TabIndex        =   188
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox Text24 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   10680
         TabIndex        =   187
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   186
         Top             =   600
         Width           =   1425
      End
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   185
         Top             =   600
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo DcbEmp 
         Height          =   315
         Left            =   6600
         TabIndex        =   189
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbSharer 
         Height          =   315
         Left            =   6600
         TabIndex        =   190
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   600
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbInves4 
         Height          =   315
         Left            =   120
         TabIndex        =   191
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì «·„»Ì⁄«  „‰"
         Height          =   285
         Index           =   28
         Left            =   11790
         RightToLeft     =   -1  'True
         TabIndex        =   204
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   26
         Left            =   8370
         RightToLeft     =   -1  'True
         TabIndex        =   203
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·—»Õ «·ﬁ«»· ·· Ê—“Ì⁄ „‰ "
         Height          =   285
         Index           =   23
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   200
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   22
         Left            =   1290
         RightToLeft     =   -1  'True
         TabIndex        =   199
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ«∆„ »«·»Ì⁄"
         Height          =   285
         Index           =   27
         Left            =   11940
         RightToLeft     =   -1  'True
         TabIndex        =   196
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„”«Â„"
         Height          =   285
         Index           =   25
         Left            =   11790
         RightToLeft     =   -1  'True
         TabIndex        =   195
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„”«Â„…"
         Height          =   285
         Index           =   24
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   194
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·„”«Â„… „‰ "
         Height          =   285
         Index           =   21
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   193
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   20
         Left            =   1290
         RightToLeft     =   -1  'True
         TabIndex        =   192
         Top             =   600
         Width           =   1890
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   161
      Top             =   4440
      Width           =   13425
      Begin VB.ComboBox DcbTypeSales2 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   183
         Top             =   960
         Width           =   1545
      End
      Begin VB.TextBox Text21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   181
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox TxtPart11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10470
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   177
         Top             =   960
         Width           =   1545
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7830
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   175
         Top             =   960
         Width           =   1425
      End
      Begin VB.TextBox Text18 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8190
         TabIndex        =   170
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3150
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   167
         Top             =   600
         Width           =   705
      End
      Begin VB.TextBox Text16 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   8190
         TabIndex        =   163
         Top             =   240
         Width           =   1065
      End
      Begin VB.ComboBox DcbTypeSales 
         Height          =   315
         Left            =   10470
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   162
         Top             =   240
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo DcbSeler 
         Height          =   315
         Left            =   4800
         TabIndex        =   164
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbLand1 
         Height          =   315
         Left            =   120
         TabIndex        =   168
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   600
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCust2 
         Height          =   315
         Left            =   4800
         TabIndex        =   171
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   600
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbInves3 
         Height          =   315
         Left            =   120
         TabIndex        =   173
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCustomerType 
         Height          =   315
         Left            =   10470
         TabIndex        =   178
         Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
         Top             =   600
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   19
         Left            =   6090
         RightToLeft     =   -1  'True
         TabIndex        =   182
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— «·„ — „‰ "
         Height          =   285
         Index           =   18
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   180
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·⁄„Ì·"
         Height          =   285
         Index           =   17
         Left            =   11790
         RightToLeft     =   -1  'True
         TabIndex        =   179
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ﬁÿ⁄Â"
         Height          =   285
         Index           =   15
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   176
         Top             =   960
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„”«Â„…"
         Height          =   285
         Index           =   14
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   174
         Top             =   240
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„Ì·"
         Height          =   285
         Index           =   13
         Left            =   8910
         RightToLeft     =   -1  'True
         TabIndex        =   172
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«—÷"
         Height          =   285
         Index           =   12
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   169
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·»«∆⁄"
         Height          =   285
         Index           =   50
         Left            =   11790
         TabIndex        =   166
         Top             =   300
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»«∆⁄"
         Height          =   285
         Index           =   11
         Left            =   9060
         RightToLeft     =   -1  'True
         TabIndex        =   165
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   146
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   149
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   148
         Top             =   240
         Width           =   3555
      End
      Begin VB.TextBox TxtNameE 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   147
         Top             =   600
         Width           =   3555
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   1200
         TabIndex        =   156
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   315
         Left            =   1200
         TabIndex        =   157
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   600
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbTypwInvse 
         Height          =   315
         Left            =   1200
         TabIndex        =   158
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   960
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbGroupInvs 
         Height          =   315
         Left            =   8280
         TabIndex        =   159
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   960
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   285
         Index           =   49
         Left            =   6840
         TabIndex        =   155
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ⁄—»Ì"
         Height          =   285
         Index           =   48
         Left            =   12000
         TabIndex        =   154
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ≈‰Ã·Ì“Ì"
         Height          =   285
         Index           =   47
         Left            =   12000
         TabIndex        =   153
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ã„Ê⁄… «·„”«Â„…"
         Height          =   285
         Index           =   46
         Left            =   12000
         TabIndex        =   152
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„”«Â„…"
         Height          =   285
         Index           =   45
         Left            =   6840
         TabIndex        =   151
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   20
         Left            =   6840
         TabIndex        =   150
         Top             =   600
         Width           =   1365
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   129
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox TxtSharValueTo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   144
         Top             =   960
         Width           =   1545
      End
      Begin VB.TextBox TxtSharNoto 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   142
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox TxtSharNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   139
         Top             =   600
         Width           =   1545
      End
      Begin VB.TextBox TxtSharValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   138
         Top             =   960
         Width           =   1545
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3960
         TabIndex        =   135
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   11010
         TabIndex        =   132
         Top             =   600
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo DcbInves 
         Height          =   315
         Left            =   7440
         TabIndex        =   130
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbSales 
         Height          =   315
         Left            =   7440
         TabIndex        =   133
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   600
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCusto 
         Height          =   315
         Left            =   240
         TabIndex        =   136
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   18
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   145
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï"
         Height          =   285
         Index           =   17
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   143
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«”Â„ „‰"
         Height          =   285
         Index           =   7
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   600
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— «·”Â„"
         Height          =   285
         Index           =   0
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   140
         Top             =   960
         Width           =   1755
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„Ì·"
         Height          =   285
         Index           =   10
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   137
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ«∆„ »«·»Ì⁄"
         Height          =   285
         Index           =   1
         Left            =   11880
         TabIndex        =   134
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·„”«Â„…"
         Height          =   285
         Index           =   16
         Left            =   11820
         RightToLeft     =   -1  'True
         TabIndex        =   131
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   108
      Top             =   4440
      Width           =   13425
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   -240
         TabIndex        =   121
         Top             =   120
         Visible         =   0   'False
         Width           =   6855
         Begin VB.TextBox TxtPartNo 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   120
            TabIndex        =   127
            Top             =   600
            Width           =   1665
         End
         Begin VB.TextBox TxtBlockNo 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3210
            TabIndex        =   125
            Top             =   600
            Width           =   1545
         End
         Begin MSDataListLib.DataCombo DcbDiv 
            Height          =   315
            Left            =   120
            TabIndex        =   122
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   240
            Width           =   4635
            _ExtentX        =   8176
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " —ﬁ„ «·ﬁÿ⁄…"
            Height          =   285
            Index           =   9
            Left            =   1410
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   600
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " —ﬁ„ «·»·Êﬂ"
            Height          =   285
            Index           =   8
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   600
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «· ﬁ”Ì„"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   5
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   240
            Width           =   1890
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         TabIndex        =   115
         Top             =   120
         Visible         =   0   'False
         Width           =   6855
         Begin VB.TextBox Text13 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3750
            TabIndex        =   116
            Top             =   240
            Width           =   1065
         End
         Begin MSDataListLib.DataCombo DcbCus 
            Height          =   315
            Left            =   120
            TabIndex        =   117
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   240
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DbDevlop 
            Height          =   315
            Left            =   120
            TabIndex        =   118
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   600
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁ«∆„ »«· ÿÊÌ—"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   240
            Width           =   1890
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " ‰Ê⁄ «· ÿÊÌ—"
            Height          =   285
            Index           =   4
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   600
            Width           =   1890
         End
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11010
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   109
         Top             =   600
         Width           =   1065
      End
      Begin MSDataListLib.DataCombo DcbInvise 
         Height          =   315
         Left            =   7440
         TabIndex        =   110
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   240
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbLand 
         Height          =   315
         Left            =   7440
         TabIndex        =   111
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   600
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«—÷"
         Height          =   285
         Index           =   6
         Left            =   11820
         RightToLeft     =   -1  'True
         TabIndex        =   113
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·„”«Â„…"
         Height          =   285
         Index           =   7
         Left            =   11820
         RightToLeft     =   -1  'True
         TabIndex        =   112
         Top             =   240
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   98
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11070
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox NameTxt 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox TxtArea 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TxtNo_planned 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox TxtPropertyDeed 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   1695
      End
      Begin VB.ComboBox DcbPaymentType 
         Height          =   315
         ItemData        =   "FrmSearchinvestment.frx":68D4
         Left            =   6960
         List            =   "FrmSearchinvestment.frx":68D6
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox TxtMeterValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   1695
      End
      Begin MSDataListLib.DataCombo DcbOwner 
         Height          =   315
         Left            =   6960
         TabIndex        =   7
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   600
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·„«·ﬂ"
         Height          =   285
         Index           =   0
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ê’› ≈‰Ã·Ì“Ì"
         Height          =   285
         Index           =   38
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ê’› ⁄—»Ì"
         Height          =   285
         Index           =   39
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„”«Õ… "
         Height          =   285
         Index           =   40
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„Œÿÿ"
         Height          =   285
         Index           =   41
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "’ﬂ «·„·ﬂÌ…"
         Height          =   285
         Index           =   42
         Left            =   12120
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·œ›⁄"
         Height          =   285
         Index           =   43
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   990
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "”⁄— «·„ —"
         Height          =   285
         Index           =   44
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   960
         Width           =   1515
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   80
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11190
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   600
         Width           =   705
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   600
         Width           =   6375
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   2280
            TabIndex        =   85
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   106692611
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   106692611
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï "
            Height          =   315
            Index           =   34
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ "
            Height          =   315
            Index           =   33
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   240
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "· «—ÌŒ"
            Height          =   195
            Index           =   31
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   240
            Width           =   1425
         End
      End
      Begin MSDataListLib.DataCombo DcbBranch10 
         Bindings        =   "FrmSearchinvestment.frx":68D8
         Height          =   315
         Left            =   7080
         TabIndex        =   81
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbEmp10 
         Bindings        =   "FrmSearchinvestment.frx":68ED
         Height          =   315
         Left            =   7080
         TabIndex        =   91
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Bindings        =   "FrmSearchinvestment.frx":6902
         Height          =   315
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbJob 
         Bindings        =   "FrmSearchinvestment.frx":6917
         Height          =   315
         Left            =   7080
         TabIndex        =   95
         Top             =   960
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…"
         Height          =   285
         Index           =   37
         Left            =   11880
         TabIndex        =   96
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   285
         Index           =   36
         Left            =   11880
         TabIndex        =   94
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«œ«—…"
         Height          =   285
         Index           =   35
         Left            =   6120
         TabIndex        =   93
         Top             =   240
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   32
         Left            =   11880
         TabIndex        =   82
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11190
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11190
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   600
         Width           =   705
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   960
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo DcbBranch6 
         Bindings        =   "FrmSearchinvestment.frx":692C
         Height          =   315
         Left            =   7080
         TabIndex        =   63
         Top             =   240
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbManger 
         Bindings        =   "FrmSearchinvestment.frx":6941
         Height          =   315
         Left            =   3960
         TabIndex        =   69
         Top             =   600
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dcsupplier 
         Height          =   315
         Left            =   3960
         TabIndex        =   73
         Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
         Top             =   960
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·„«·ﬂ"
         Height          =   285
         Index           =   1
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ— «·„”«Â„…"
         Height          =   285
         Index           =   29
         Left            =   11880
         TabIndex        =   78
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   30
         Left            =   11880
         TabIndex        =   76
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„”«Â„…"
         Height          =   285
         Index           =   28
         Left            =   5760
         TabIndex        =   74
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«”Â„"
         Height          =   285
         Index           =   27
         Left            =   2280
         TabIndex        =   70
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·„”«Â„… «·«Ã„«·Ì…"
         Height          =   285
         Index           =   26
         Left            =   2280
         TabIndex        =   66
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·”Â„"
         Height          =   285
         Index           =   25
         Left            =   2280
         TabIndex        =   64
         Top             =   960
         Width           =   1605
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox TxtShareValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox TxtInvesTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TxtInvesNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DcbBranch1 
         Bindings        =   "FrmSearchinvestment.frx":6956
         Height          =   315
         Left            =   8040
         TabIndex        =   50
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbTypwInvse1 
         Bindings        =   "FrmSearchinvestment.frx":696B
         Height          =   315
         Left            =   8040
         TabIndex        =   52
         Top             =   600
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·”Â„"
         Height          =   285
         Index           =   23
         Left            =   2280
         TabIndex        =   61
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·„”«Â„… «·«Ã„«·Ì…"
         Height          =   285
         Index           =   22
         Left            =   2280
         TabIndex        =   59
         Top             =   240
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«”Â„"
         Height          =   285
         Index           =   21
         Left            =   6480
         TabIndex        =   57
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·„”«Â„…"
         Height          =   285
         Index           =   19
         Left            =   6480
         TabIndex        =   55
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„”«Â„…"
         Height          =   285
         Index           =   16
         Left            =   11880
         TabIndex        =   53
         Top             =   600
         Width           =   1605
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   24
         Left            =   11880
         TabIndex        =   51
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   4440
      Width           =   13425
      Begin VB.TextBox TxtShareInvsCount 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox TxtToShareValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtCountShare 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   11070
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   720
         Width           =   825
      End
      Begin VB.TextBox TxtFromShareValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   240
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Bindings        =   "FrmSearchinvestment.frx":6980
         Height          =   315
         Left            =   8040
         TabIndex        =   34
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo DcbSahr 
         Bindings        =   "FrmSearchinvestment.frx":6995
         Height          =   315
         Left            =   4320
         TabIndex        =   43
         Top             =   720
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«”Â„"
         Height          =   285
         Index           =   12
         Left            =   2880
         TabIndex        =   44
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·„”«Â„"
         Height          =   285
         Index           =   15
         Left            =   12000
         TabIndex        =   41
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   11
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·”Â„ „‰"
         Height          =   285
         Index           =   10
         Left            =   2880
         TabIndex        =   38
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«”Â„ «·„ «Õ…"
         Height          =   285
         Index           =   9
         Left            =   6480
         TabIndex        =   36
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   285
         Index           =   8
         Left            =   11880
         TabIndex        =   35
         Top             =   240
         Width           =   1605
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   0
      Width           =   13665
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ «·„”«Â„« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   5400
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   12360
         Picture         =   "FrmSearchinvestment.frx":69AA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   3015
      Left            =   0
      TabIndex        =   29
      Top             =   720
      Width           =   13455
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid14 
         Height          =   2835
         Left            =   -30
         TabIndex        =   280
         Top             =   60
         Visible         =   0   'False
         Width           =   13395
         _cx             =   23627
         _cy             =   5001
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":15299
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2745
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4842
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":1550C
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
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2865
         Left            =   120
         TabIndex        =   30
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   5054
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":1565E
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
         Height          =   2625
         Left            =   120
         TabIndex        =   48
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":157CC
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid6 
         Height          =   2625
         Left            =   120
         TabIndex        =   77
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":15948
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
         Height          =   2625
         Left            =   120
         TabIndex        =   83
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":15AE7
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
         Height          =   2625
         Left            =   120
         TabIndex        =   97
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":15C74
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid5 
         Height          =   2625
         Left            =   120
         TabIndex        =   107
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":15E25
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid7 
         Height          =   2625
         Left            =   120
         TabIndex        =   114
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":15FA8
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid8 
         Height          =   2625
         Left            =   120
         TabIndex        =   128
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":1614A
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid9 
         Height          =   2625
         Left            =   120
         TabIndex        =   160
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":162FB
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid10 
         Height          =   2625
         Left            =   120
         TabIndex        =   205
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   4630
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":164A2
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid11 
         Height          =   2865
         Left            =   120
         TabIndex        =   219
         Top             =   120
         Width           =   13155
         _cx             =   23204
         _cy             =   5054
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   16777088
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchinvestment.frx":16623
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
      BackColor       =   &H00E2E9E9&
      Height          =   615
      Left            =   0
      TabIndex        =   25
      Top             =   5880
      Width           =   13455
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   10
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label lblL 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   2145
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·≈Ã„«·Ì"
         Height          =   285
         Index           =   2
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   21
      Top             =   6480
      Width           =   13455
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   9480
         TabIndex        =   22
         Top             =   240
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   661
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
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
         ButtonImage     =   "FrmSearchinvestment.frx":16748
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
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   23
         Top             =   240
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   661
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
         ButtonImage     =   "FrmSearchinvestment.frx":1CFAA
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
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   24
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
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
         ButtonImage     =   "FrmSearchinvestment.frx":2380C
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
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3720
      Width           =   7035
      Begin VB.TextBox TxtIDTO 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   1635
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·Õ—ﬂ…"
         Height          =   195
         Index           =   14
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   315
         Index           =   6
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   5
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3720
      Width           =   6375
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2280
         TabIndex        =   2
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107413507
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107413507
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Õ—ﬂ…"
         Height          =   195
         Index           =   13
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ "
         Height          =   315
         Index           =   4
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï "
         Height          =   315
         Index           =   3
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   1080
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid13 
      Height          =   2835
      Left            =   120
      TabIndex        =   231
      Top             =   840
      Visible         =   0   'False
      Width           =   13155
      _cx             =   23204
      _cy             =   5001
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777088
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchinvestment.frx":4D42E
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
   Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid12 
      Height          =   2865
      Left            =   120
      TabIndex        =   230
      Top             =   840
      Width           =   13155
      _cx             =   23204
      _cy             =   5054
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777088
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchinvestment.frx":4D5E3
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
Attribute VB_Name = "FrmSearchinvestment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
'Dim DCboSearch As FrmGeneralFundReceipt
Public inde As Integer


Private Sub ChCarType_Click(index As Integer)
If ChCarType(0).value = True Then
DcbSupplem.Enabled = True
DCCar.Enabled = True
DcbSupplem2.Enabled = False
DcbCar2.BoundText = ""
Else
DcbSupplem2.Enabled = True
DcbCar2.Enabled = True
DCCar.BoundText = ""
DCCar.Enabled = False
DcbSupplem.BoundText = ""
DcbSupplem.Enabled = False
End If
End Sub

Private Sub DcbCar2_Change()
DcbCar2_Click (0)
End Sub

Private Sub DcbCar2_Click(Area As Integer)
 Dim Dcombos As New ClsDataCombos
Dcombos.GetBartCarByVonder DcbSupplem2, val(DcbCar2.BoundText)
End Sub

Private Sub DcbCus_Change()
If val(DcbCus.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCus.BoundText, EmpCode
    Text13.text = EmpCode
End Sub

Private Sub DcbCus_Click(Area As Integer)
DcbCus_Change
End Sub

Private Sub DcbCust2_Change()
  If val(DcbCust2.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCust2.BoundText, EmpCode
    Me.Text18.text = EmpCode
End Sub

Private Sub DcbCust2_Click(Area As Integer)
DcbCust2_Change
End Sub

Private Sub DcbCusto_Change()
  If val(DcbCusto.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCusto.BoundText, EmpCode
    Me.Text15.text = EmpCode
End Sub

Private Sub DcbCusto_Click(Area As Integer)
DcbCusto_Change
End Sub

Private Sub DcbEmp_Change()
 If val(DcbEmp.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmp.BoundText, EmpCode
    Me.Text26.text = EmpCode
End Sub

Private Sub DcbEmp_Click(Area As Integer)
DcbEmp_Change
End Sub

Private Sub DcbEmp10_Change()
DcbEmp10_Click (0)
End Sub

Private Sub DcbEmp10_Click(Area As Integer)
 If val(DcbEmp10.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmp10.BoundText, EmpCode
    Text8.text = EmpCode
    
End Sub

Private Sub DcbLand_Change()
Dim Fullcode As String
If val(DcbLand.BoundText) <> 0 Then
GetTblBuyLandRealEstate val(DcbLand.BoundText), Fullcode, 0
Me.Text12.text = Fullcode
End If
End Sub

Private Sub DcbLand_Click(Area As Integer)
DcbLand_Change
End Sub

Private Sub DcbLand1_Change()
Dim Fullcode As String
If val(DcbLand1.BoundText) <> 0 Then
GetTblBuyLandRealEstate val(DcbLand1.BoundText), Fullcode, 0
Me.Text17.text = Fullcode
End If
End Sub

Private Sub DcbLand1_Click(Area As Integer)
DcbLand1_Change
End Sub

Private Sub DcbManger_Change()
DcbManger_Click (0)
End Sub

Private Sub DcbManger_Click(Area As Integer)
If val(DcbManger.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbManger.BoundText, EmpCode
    Me.Text7.text = EmpCode
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
End Sub

Private Sub DcbOwner_Change()
   If val(DcbOwner.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbOwner.BoundText, EmpCode
    Me.Text9.text = EmpCode
End Sub

Private Sub DcbSahr_Change()
DcbSahr_Click (0)
End Sub

Private Sub DcbSahr_Click(Area As Integer)
Dim Fullcode As String
      Fullcode = ""
        GetCustomersDetail val(Me.DcbSahr.BoundText), , Fullcode
        Text2.text = Fullcode
End Sub

Private Sub DcbSahr_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 22
            FrmCustemerSearch.show vbModal
            
        End If
End Sub

Private Sub DcbSales_Change()
  If val(DcbSales.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbSales.BoundText, EmpCode
    Me.Text14.text = EmpCode
End Sub

Private Sub DcbSales_Click(Area As Integer)
DcbSales_Change
End Sub

Private Sub DcbSharer_Change()
  If val(DcbSharer.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbSharer.BoundText, EmpCode
    Me.Text24.text = EmpCode
End Sub

Private Sub DcbSharer_Click(Area As Integer)
DcbSharer_Change
End Sub

Private Sub DcbTypeSales_Change()
Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos
            Dcombos.ClearMyDataCombo Me.DcbSales
    Select Case DcbTypeSales.ListIndex
        Case 0
            Set Dcombos = New ClsDataCombos
            Dcombos.GetEmployees DcbSeler
          Case 1
          Set Dcombos = New ClsDataCombos
            Dcombos.GetCustomersSuppliers 2, Me.DcbSeler, True
          End Select
End Sub

Private Sub DcbTypeSales_Click()
DcbTypeSales_Change
End Sub

Private Sub dcCar_Change()
dcCar_Click (0)
End Sub

Private Sub dcCar_Click(Area As Integer)
     Dim Dcombos As New ClsDataCombos
Dcombos.GetPartCar DcbSupplem, val(DCCar.BoundText)
End Sub

Private Sub FG_Click()
If inde = 0 Then
Frminvestment.FindRec val(Fg.TextMatrix(Fg.row, 1))
ElseIf inde = 1 Then
FrmIPOSharer.TxtOrderInvse.text = val(Fg.TextMatrix(Fg.row, 1))
ElseIf inde = 2 Then
FrmIPO.TxtOrderInvse.text = val(Fg.TextMatrix(Fg.row, 1))
ElseIf inde = 5 Then
FrmActiveInvestment.TxtInviseOrder.text = val(Fg.TextMatrix(Fg.row, 1))
End If
End Sub
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim GrdBack As ClsBackGroundPic
    Dim My_SQL As String
    If SystemOptions.UserInterface = ArabicInterface Then
        With DcbImportExport
            .Clear
            .AddItem "’«œ—"
            .AddItem "Ê«—œ"
        End With
    With DcbPaymentType
    .Clear
    .AddItem "‰ﬁœÌ"
    .AddItem "«Ã·"
    End With
  With DcbTypeSales
 .Clear
 .AddItem "„ÊŸ›"
 .AddItem "„Ê—œ"
 End With
   With DcbTypeSales2
 .Clear
 .AddItem "„ÊŸ›"
 .AddItem "„Ê—œ"
 End With
    Else
        With DcbImportExport
            .Clear
            .AddItem "Import"
            .AddItem "Export"
        End With
       With DcbPaymentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With
  With DcbTypeSales
 .Clear
 .AddItem "Employee"
 .AddItem "Vendor"
 End With
   With DcbTypeSales2
 .Clear
 .AddItem "Employee"
 .AddItem "Vendor"
 End With
    End If
    
    
     If SystemOptions.IsHiddenTransportInv Then
        
        VSFlexGrid14.TextMatrix(0, VSFlexGrid14.ColIndex("ManualNo")) = "—ﬁ„ ÿ·» «—«„ﬂÊ"
    End If
    
    Frame16.Visible = False
    VSFlexGrid11.Visible = False
    Frame15.Visible = False
    VSFlexGrid10.Visible = False
    Frame14.Visible = False
    VSFlexGrid9.Visible = False
    VSFlexGrid8.Visible = False
     Frame13.Visible = False
    Frame11.Visible = False
    Frame12.Visible = False
    VSFlexGrid7.Visible = False
    Frame10.Visible = False
    VSFlexGrid5.Visible = False
    VSFlexGrid4.Visible = False
    Frame9.Visible = False
    VSFlexGrid2.Visible = False
    Frame7.Visible = False
    Frame5.Visible = False
      VSFlexGrid1.Visible = False
      Frame4.Visible = False
       VSFlexGrid6.Visible = False
      Frame6.Visible = False
      Fg.Visible = False
       lblLW.Visible = False
      VSFlexGrid3.Visible = False
      VSFlexGrid14.Visible = False
      Frame22.Visible = False
If inde = 3 Then
      VSFlexGrid1.Visible = True
      Frame4.Visible = True
      Label1(2).Caption = "»ÕÀ ≈ﬂ  «» «·„”«Â„Ì‰"
ElseIf inde = 4 Then
      VSFlexGrid2.Visible = True
     Frame5.Visible = True
      Label1(2).Caption = "»ÕÀ › Õ «·«ﬂ  «» ›Ì «·„”«Â„… "
ElseIf inde = 10 Then
      VSFlexGrid3.Visible = True
      Frame7.Visible = True
      Label1(2).Caption = "»ÕÀ  ﬁ—Ì— ”Ì— ⁄„· ÌÊ„Ì"
ElseIf inde = 6 Or inde = 16 Or inde = 17 Or inde = 18 Or inde = 19 Or inde = 20 Or inde = 28 Then
      VSFlexGrid6.Visible = True
      Frame6.Visible = True
      Label1(2).Caption = "»ÕÀ    ›⁄Ì·  «·„”«Â„… "
ElseIf inde = 7 Or inde = 8 Or inde = 71 Then
      VSFlexGrid4.Visible = True
      Frame9.Visible = True
      Label1(2).Caption = "»ÕÀ ‘—«¡ «·«—«÷Ì Ê«·⁄ﬁ«—« "
ElseIf inde = 11 Or inde = 111 Or inde = 110 Then
      Frame10.Visible = True
      Frame11.Visible = True
    VSFlexGrid5.Visible = True
    If inde = 11 Or inde = 110 Then
    Label1(2).Caption = "»ÕÀ „’—Ê›«  «· ÿÊÌ— "
   Else
    Label1(2).Caption = "»ÕÀ „—œÊœ«  „’—Ê›«  «· ÿÊÌ— "
   End If
ElseIf inde = 12 Then
      Frame10.Visible = True
    Frame12.Visible = True
    VSFlexGrid7.Visible = True
    Label1(2).Caption = "»ÕÀ  ﬁ”Ì„ «·«—«÷Ì "
ElseIf inde = 13 Then
      VSFlexGrid8.Visible = True
    Frame13.Visible = True
    Label1(2).Caption = "»ÕÀ  ≈‘⁄«—  ‰«“·/»Ì⁄ «”Â„ "
ElseIf inde = 14 Or inde = 72 Then
   VSFlexGrid9.Visible = True
    Frame14.Visible = True
    Label1(2).Caption = "»ÕÀ  ›« Ê—… „»Ì⁄«  "
ElseIf inde = 15 Then
   VSFlexGrid10.Visible = True
    Frame15.Visible = True
    Label1(2).Caption = "»ÕÀ ≈À»«  «—»«Õ «·„”«Â„Ì‰"
ElseIf inde = 27 Then
    Frame16.Visible = True
    VSFlexGrid12.Visible = True
    Label1(2).Caption = "»ÕÀ ”‰œ«  ÕÃ“ «·ﬁÿ⁄"
'############### khaled #######################
ElseIf inde = 30 Then
    Frame18.Visible = True
    VSFlexGrid11.Visible = True
    Label1(2).Caption = "»ÕÀ ”‰œ«  ÕÃ“ «·ÊÕœ« "
ElseIf inde = 31 Then
    Frame19.Visible = True
    VSFlexGrid13.Visible = True
    Label1(2).Caption = "»ÕÀ «—‘Ì› «·„⁄«„·« "
    lbl(14).Caption = "—ﬁ„ «·„⁄«„·…"
'##############################################
ElseIf inde = 32 Then
    Frame22.Visible = True
    VSFlexGrid14.Visible = True
    Label1(2).Caption = "»ÕÀ »Ì«‰«  «·—Õ·« "
    lbl(14).Caption = "—ﬁ„ «·—Õ·…"
'##############################################
Else
      Label1(2).Caption = "»ÕÀ «·„”«Â„« "
      Fg.Visible = True
      lblLW.Visible = True
End If
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
   '############### khaled #######################
   Dcombos.GetProjecInvestment Me.DcbProject
   Dcombos.GetBranches Me.DcbKBranch
   Dcombos.GetBranches ArchSearchBranchDC
   '##############################################
    Dcombos.GetCustomersSuppliers 2, Me.DcbOwner
    Dcombos.GetEmployees Me.DcbEmployee
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcbEmp10
    Dcombos.GetEmployees Me.DcbManger
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetBranches Me.DcbBranch
    Dcombos.GetBranches Me.DcbBranch1
    Dcombos.GetBranches Me.DcbBranch6
    Dcombos.GetBranches Me.DcbBranch10
    Dcombos.GetInvestmentType Me.DcbTypwInvse
    Dcombos.GetInvestmentGroup Me.DcbGroupInvs
   Dcombos.GetInvestmentType Me.DcbTypwInvse1
   Dcombos.GetCustomersSuppliers 2, Me.dcsupplier
   Dcombos.GetEmpDepartments Me.DcbDept
   Dcombos.GetEmpJobsTypes Me.DcbJob
   Dcombos.GetBuyLandRealEstate DcbLand
   Dcombos.GetInvestmentActive Me.DcbInvise
   Dcombos.GetInvestmentActive Me.DcbInvest
   Dcombos.GetInvestmentActive Me.DcbInves
   Dcombos.GetInvestmentActive Me.DcbInves3
   Dcombos.GetInvestmentActive Me.DcbInves4
   Dcombos.GetBuyLandRealEstate DcbLand1
   Dcombos.GetEmployees Me.DcbEmp
   Dcombos.GetInvStoreType Me.DcCustomerType
      Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetShips Me.DcbShip
    Dcombos.GetCitiesDistance Me.DcCityFromId, 0
    Dcombos.GetCitiesDistance Me.DcCityToId, 1
    Dcombos.GetTblCarsDataGroup VehicleType
    Dcombos.GetEmployees Me.DCEmp, , True
    Dcombos.GetCars Me.DCCar
    Dcombos.GetCarByVonder DcbCar2
    Dcombos.GetBranches DcbBranch22
    
     If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=20 and Flg=1  order by CusName"
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=20 and Flg=1  order by CusNamee"
    End If
    fill_combo DcbSales, My_SQL
    fill_combo DcbCusto, My_SQL
    fill_combo DcbCust2, My_SQL
    fill_combo DcbSharer, My_SQL
      If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=20 and Flg=1  order by CusName"
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=20 and Flg=1  order by CusNamee"
    End If
    fill_combo DcbSahr, My_SQL
          If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select ID,Name from TblSpreading    order by Name"
    Else
        My_SQL = "  select ID,NameE from TblSpreading    order by NameE"
    End If
    fill_combo DcbDiv, My_SQL
    
     If SystemOptions.UserInterface = ArabicInterface Then
                   My_SQL = "  select CusID,CusName from TblCustemers  where type=2  order by CusName"
     Else
                   My_SQL = "  select CusID,CusNamee from TblCustemers  where type=2  order by CusNamee"
    End If
       fill_combo DcbCus, My_SQL
          If SystemOptions.UserInterface = ArabicInterface Then
                   My_SQL = "  select ID,Name from TblInvestmentsGroup  order by Name"
     Else
                   My_SQL = "  select ID,NameE from TblInvestmentsGroup   order by NameE"
    End If
       fill_combo DbDevlop, My_SQL
       
      Set GrdBack = New ClsBackGroundPic
    With Me.Fg
        Set .WallPaper = GrdBack.Picture
       .AutoSize 0, .Cols - 1, False
    End With
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
     SetDtpickerDate Me.DtpDateFrom
     SetDtpickerDate Me.DtpDateTo
     SetDtpickerDate Me.DTPicker1
     SetDtpickerDate Me.DTPicker2
     SetDtpickerDate FromLiqDate
     SetDtpickerDate ToLiqDate
     FrmEnterDate.value = Now
     ToEnterDate.value = Now
     SetDtpickerDate DtpDateFrom
     SetDtpickerDate DtpDateTo
     FrmExitDate.value = Now
     ToExitDate.value = Now
      FrmExitDate.value = ""
     ToExitDate.value = ""
     FrmEnterDate.value = ""
     ToEnterDate.value = ""
     
   End Sub
   Private Sub Cmd_Click(index As Integer)
    Select Case index
    
        Case 0
        If inde = 3 Then
        GetDataShare
        ElseIf inde = 4 Then
        GetDataIPO
        ElseIf inde = 6 Or inde = 16 Or inde = 17 Or inde = 18 Or inde = 19 Or inde = 20 Or inde = 28 Then
        GetDataActiveInvest
        ElseIf inde = 7 Or inde = 8 Or inde = 71 Then
        GetDataBuyLandRealEstate
        ElseIf inde = 10 Then
        GetDataWorkFlow
        ElseIf inde = 11 Or inde = 111 Or inde = 110 Then
        If inde = 11 Or inde = 110 Then
        GetDataExpensesLand
        ElseIf inde = 111 Then
        GetDataExpensesLand2
        End If
        ElseIf inde = 12 Then
        GetDataDivideLand
         ElseIf inde = 13 Then
        GetDataBuyBillInvestment
        ElseIf inde = 14 Or inde = 72 Then
        GetDataSalesBill
        ElseIf inde = 15 Then
        GetDataInvestProfit
         ElseIf inde = 27 Then
        GetDataInvestLiquidation
        '############## Khaled #################
        ElseIf inde = 30 Then
            GetKData
        ElseIf inde = 31 Then
            GetArchData
        ElseIf inde = 32 Then
            GetTripData
        '#######################################
        Else
        GetData
        End If
        Case 1
        clear_all Me
        ToLiqDate.value = ""
        FromLiqDate.value = ""
          Me.DtpDateFrom.value = ""
          Me.DtpDateTo.value = ""
          Me.DTPicker1.value = ""
          Me.DTPicker2.value = ""
          FrmEnterDate.value = ""
          ToEnterDate.value = ""
          DtpDateFrom.value = ""
          DtpDateTo.value = ""
          FrmExitDate.value = ""
          ToExitDate.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lblL(0).Caption = "Search Results"
            End If
            Case 2
            Unload Me
    End Select
End Sub
Private Sub Form_Activate()
 '   PutFormOnTop Me.hWnd
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    'Set DCboSearch = Nothing
End Sub
Public Sub GetDataActiveInvest()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblActivateInvestment.ID, dbo.TblActivateInvestment.RecorDate, dbo.TblActivateInvestment.TypePrstg, dbo.TblActivateInvestment.PercenValue, "
    sql = sql & "                  dbo.TblActivateInvestment.InviseOrder, dbo.TblActivateInvestment.InviseNo, dbo.TblActivateInvestment.InviseValue, dbo.TblActivateInvestment.SharesCount,"
    sql = sql & "                  dbo.TblActivateInvestment.SharesValue, dbo.TblActivateInvestment.Typ, dbo.TblActivateInvestment.Description, dbo.TblActivateInvestment.Remarks,"
    sql = sql & "                  dbo.TblActivateInvestment.Area, dbo.TblActivateInvestment.No_planned, dbo.TblActivateInvestment.MeterValue, dbo.TblActivateInvestment.Identityof,"
    sql = sql & "                  dbo.TblActivateInvestment.Address, dbo.TblActivateInvestment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
    sql = sql & "                  dbo.TblActivateInvestment.InvManager, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
    sql = sql & "                  dbo.TblActivateInvestment.OwnerID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode"
    sql = sql & "      FROM         dbo.TblActivateInvestment LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblCustemers ON dbo.TblActivateInvestment.OwnerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee ON dbo.TblActivateInvestment.InvManager = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblActivateInvestment.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = ""
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblActivateInvestment.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblActivateInvestment.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblActivateInvestment.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblActivateInvestment.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblActivateInvestment.RecorDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblActivateInvestment.RecorDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblActivateInvestment.RecorDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblActivateInvestment.RecorDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
        If Me.DcbBranch6.text <> "" And (val(DcbBranch6.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblActivateInvestment.BranchID =" & Me.DcbBranch6.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblActivateInvestment.BranchID =" & Me.DcbBranch6.BoundText & ""
       End If
     End If
         If Me.DcbManger.text <> "" And (val(DcbManger.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblActivateInvestment.InvManager =" & Me.DcbManger.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblActivateInvestment.InvManager =" & Me.DcbManger.BoundText & ""
       End If
     End If
        If Me.dcsupplier.text <> "" And (val(dcsupplier.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblActivateInvestment.OwnerID =" & Me.dcsupplier.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblActivateInvestment.OwnerID =" & Me.dcsupplier.BoundText & ""
       End If
     End If
     
       If Me.Text6.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblActivateInvestment.InviseNo =" & val(Me.Text6.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblActivateInvestment.InviseNo =" & (Me.Text6.text) & ""
       End If
     End If
      If Me.Text4.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblActivateInvestment.InviseValue =" & val(Me.Text4.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblActivateInvestment.InviseValue =" & val(Me.Text4.text) & ""
       End If
     End If
     
      If Me.Text5.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblActivateInvestment.SharesCount =" & val(Me.Text5.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblActivateInvestment.SharesCount =" & val(Me.Text5.text) & ""
       End If
     End If
     
     If Me.Text3.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblActivateInvestment.SharesValue =" & val(Me.Text3.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblActivateInvestment.SharesValue =" & val(Me.Text3.text) & ""
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblActivateInvestment.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid6
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecorDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecorDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InviseNo")) = IIf(IsNull(rs("InviseNo").value), 0, rs("InviseNo").value)
               .TextMatrix(i, .ColIndex("InviseValue")) = IIf(IsNull(rs("InviseValue").value), 0, rs("InviseValue").value)
                .TextMatrix(i, .ColIndex("SharesCount")) = IIf(IsNull(rs("SharesCount").value), 0, rs("SharesCount").value)
                .TextMatrix(i, .ColIndex("SharesValue")) = IIf(IsNull(rs("SharesValue").value), 0, rs("SharesValue").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataWorkFlow()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.TblDailyWorkflow.ID, dbo.TblDailyWorkflow.RecordDate, dbo.TblDailyWorkflow.NameDay, dbo.TblDailyWorkflow.EmpID, dbo.TblEmployee.Emp_Name, "
    sql = sql & "                  dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode,"
    sql = sql & "                   dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4,"
    sql = sql & "                   dbo.TblDailyWorkflow.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblDailyWorkflow.DeptID,"
    sql = sql & "                   dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblDailyWorkflow.PerformType, dbo.TblDailyWorkflow.Remarks,"
    sql = sql & "                   dbo.TblDailyWorkflow.DirectManagerID, TblEmployee_1.Emp_Name AS ManEmp_Name, TblEmployee_1.Emp_Name1 AS ManEmp_Name1,"
    sql = sql & "                   TblEmployee_1.Emp_Name2 AS ManEmp_Name2, TblEmployee_1.Emp_Name3 AS ManEmp_Name3, TblEmployee_1.Emp_Name4 AS ManEmp_Name4,"
    sql = sql & "                   TblEmployee_1.Fullcode AS ManFullcode, TblEmployee_1.Emp_Namee4 AS ManEmp_Namee4, TblEmployee_1.Emp_Namee3 AS ManEmp_Namee3,"
    sql = sql & "                   TblEmployee_1.Emp_Namee2 AS ManEmp_Namee2, TblEmployee_1.Emp_Namee1 AS ManEmp_Namee1, TblEmployee_1.Emp_Namee AS ManEmp_Namee,"
    sql = sql & "                   dbo.TblDailyWorkflow.MangerName, dbo.TblDailyWorkflow.TransDate, dbo.TblDailyWorkflow.FrmTime, dbo.TblDailyWorkflow.TOTime,"
    sql = sql & "                   dbo.TblDailyWorkflow.BranchId , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
    sql = sql & "   FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
    sql = sql & "                   dbo.TblDailyWorkflow ON dbo.TblBranchesData.branch_id = dbo.TblDailyWorkflow.BranchID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmployee TblEmployee_1 ON dbo.TblDailyWorkflow.DirectManagerID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmpDepartments ON dbo.TblDailyWorkflow.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmpJobsTypes ON dbo.TblDailyWorkflow.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmployee ON dbo.TblDailyWorkflow.EmpID = dbo.TblEmployee.Emp_ID"

       BolBegine = False
       StrWhere = ""
  
    
   ''''''''''''''''''''''''''''''''''''
    
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblDailyWorkflow.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDailyWorkflow.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblDailyWorkflow.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblDailyWorkflow.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDailyWorkflow.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDailyWorkflow.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDailyWorkflow.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblDailyWorkflow.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''5456
         If Not IsNull(Me.DTPicker1.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDailyWorkflow.TransDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDailyWorkflow.TransDate>=" & SQLDate(Me.DTPicker1.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DTPicker2.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDailyWorkflow.TransDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblDailyWorkflow.TransDate<=" & SQLDate(Me.DTPicker2.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
        If Me.DcbBranch10.text <> "" And (val(DcbBranch10.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblDailyWorkflow.BranchID =" & Me.DcbBranch10.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblDailyWorkflow.BranchID =" & Me.DcbBranch10.BoundText & ""
       End If
     End If
         If Me.DcbDept.text <> "" And (val(DcbDept.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblDailyWorkflow.DeptID =" & Me.DcbDept.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblDailyWorkflow.DeptID =" & Me.DcbDept.BoundText & ""
       End If
     End If
        If Me.DcbEmp10.text <> "" And (val(DcbEmp10.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblDailyWorkflow.EmpID =" & Me.DcbEmp10.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblDailyWorkflow.EmpID =" & Me.DcbEmp10.BoundText & ""
       End If
     End If
     
        If Me.DcbJob.text <> "" And (val(DcbJob.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblDailyWorkflow.JobID =" & Me.DcbJob.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblDailyWorkflow.JobID =" & Me.DcbJob.BoundText & ""
       End If
     End If

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblDailyWorkflow.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid3
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                  If Not (IsNull(rs("TransDate").value)) Then
                .TextMatrix(i, .ColIndex("TransDate")) = Format(rs("TransDate").value, "yyyy/M/d")
                End If
                
                .TextMatrix(i, .ColIndex("NameDay")) = IIf(IsNull(rs("NameDay").value), "", rs("NameDay").value)
               '.TextMatrix(i, .ColIndex("PerformType")) = IIf(IsNull(rs("InviseValue").value), 0, rs("InviseValue").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                
                Else
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeNamee").value), "", rs("JobTypeNamee").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataIPO()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = " SELECT     dbo.TblIPO.ID, dbo.TblIPO.RecorDate, dbo.TblIPO.Remark, dbo.TblIPO.InvesNo, dbo.TblIPO.InvesTotal, dbo.TblIPO.CountShare, dbo.TblIPO.ShareValue, "
    sql = sql & "                   dbo.TblIPO.OrderInvse, dbo.TblIPO.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblIPO.TypwInvse,"
    sql = sql & "                  dbo.TblShareType.name , dbo.TblShareType.NameE"
    sql = sql & "  FROM         dbo.TblIPO LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblShareType ON dbo.TblIPO.TypwInvse = dbo.TblShareType.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.TblIPO.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = ""
  
    
   ''''''''''''''''''''''''''''''''''''
    
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblIPO.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblIPO.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblIPO.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblIPO.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblIPO.RecorDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblIPO.RecorDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblIPO.RecorDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblIPO.RecorDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcbTypwInvse1.text <> "" And (val(DcbTypwInvse1.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPO.TypwInvse =" & Me.DcbTypwInvse1.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPO.TypwInvse =" & Me.DcbTypwInvse1.BoundText & ""
       End If
     End If
        If Me.DcbBranch1.text <> "" And (val(DcbBranch1.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPO.BranchID =" & Me.DcbBranch1.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPO.BranchID =" & Me.DcbBranch1.BoundText & ""
       End If
     End If
       If Me.TxtInvesNo.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPO.InvesNo =" & val(Me.TxtInvesNo.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPO.InvesNo =" & (Me.TxtInvesNo.text) & ""
       End If
     End If
      If Me.TxtInvesTotal.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPO.InvesTotal =" & val(Me.TxtInvesTotal.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPO.InvesTotal =" & val(Me.TxtInvesTotal.text) & ""
       End If
     End If
     
      If Me.Text1.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPO.CountShare =" & val(Me.Text1.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPO.CountShare =" & val(Me.Text1.text) & ""
       End If
     End If
     
     If Me.TxtShareValue.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPO.ShareValue =" & val(Me.TxtShareValue.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPO.ShareValue =" & val(Me.TxtShareValue.text) & ""
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblIPO.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid2
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecorDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecorDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InvesNo")) = IIf(IsNull(rs("InvesNo").value), 0, rs("InvesNo").value)
                .TextMatrix(i, .ColIndex("InvesTotal")) = IIf(IsNull(rs("InvesTotal").value), 0, rs("InvesTotal").value)
                .TextMatrix(i, .ColIndex("CountShare")) = IIf(IsNull(rs("CountShare").value), 0, rs("CountShare").value)
                .TextMatrix(i, .ColIndex("ShareValue")) = IIf(IsNull(rs("ShareValue").value), 0, rs("ShareValue").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataBuyLandRealEstate()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.TblBuyLanReEst.ID, dbo.TblBuyLanReEst.RecordDate, dbo.TblBuyLanReEst.Name, dbo.TblBuyLanReEst.NameE, dbo.TblBuyLanReEst.FullCode, "
    sql = sql & "                   dbo.TblBuyLanReEst.BranchID, dbo.TblBuyLanReEst.No_planned, dbo.TblBuyLanReEst.Area, dbo.TblBuyLanReEst.MeterPrice, dbo.TblBuyLanReEst.Total,"
    sql = sql & "                     dbo.TblBuyLanReEst.TitledeedNo, dbo.TblBuyLanReEst.PayType, dbo.TblBuyLanReEst.InstalNo, dbo.TblBuyLanReEst.FristDate, dbo.TblBuyLanReEst.Period,"
    sql = sql & "                     dbo.TblBuyLanReEst.PeriodType, dbo.TblBuyLanReEst.Remarks, dbo.TblBuyLanReEst.OwnerID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    sql = sql & "                     dbo.TblCustemers.Fullcode AS CusFullcode"
    sql = sql & "     FROM         dbo.TblBuyLanReEst LEFT OUTER JOIN"
    sql = sql & "                     dbo.TblCustemers ON dbo.TblBuyLanReEst.ID = dbo.TblCustemers.CusID"
       BolBegine = False
       StrWhere = ""
  
    
   ''''''''''''''''''''''''''''''''''''
    
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblBuyLanReEst.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyLanReEst.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblBuyLanReEst.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblBuyLanReEst.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblBuyLanReEst.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyLanReEst.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblBuyLanReEst.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblBuyLanReEst.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcbOwner.text <> "" And (val(DcbOwner.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblBuyLanReEst.OwnerID =" & Me.DcbOwner.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblBuyLanReEst.OwnerID =" & Me.DcbOwner.BoundText & ""
       End If
     End If
  If Me.DcbPaymentType.text <> "" And (val(DcbPaymentType.ListIndex) <> -1) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblBuyLanReEst.PayType =" & Me.DcbPaymentType.ListIndex & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblBuyLanReEst.PayType =" & Me.DcbPaymentType.ListIndex & ""
       End If
     End If
     
       If Me.TxtMeterValue.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblBuyLanReEst.MeterPrice =" & val(Me.TxtMeterValue.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblBuyLanReEst.MeterPrice =" & (Me.TxtMeterValue.text) & ""
       End If
     End If
      If Me.TxtArea.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblBuyLanReEst.Area =" & val(Me.TxtArea.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblBuyLanReEst.Area =" & val(Me.TxtArea.text) & ""
       End If
     End If
     
      If Me.TxtNo_planned.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblBuyLanReEst.No_planned ='" & Me.TxtNo_planned.text & "'"
        Else:
          BolBegine = True
          StrWhere = " Where TblBuyLanReEst.No_planned ='" & Me.TxtNo_planned.text & "'"
       End If
     End If
     
      If Me.TxtPropertyDeed.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblBuyLanReEst.TitledeedNo ='" & Me.TxtPropertyDeed.text & "'"
        Else:
          BolBegine = True
          StrWhere = " Where TblBuyLanReEst.TitledeedNo ='" & Me.TxtPropertyDeed.text & "'"
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 If Me.NameTxt.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblBuyLanReEst.Name like '%" & Me.NameTxt.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyLanReEst.Name  like '%" & Me.NameTxt.text & "%'"
        End If
    End If
    
     If Me.Text11.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblBuyLanReEst.NameE like '%" & Me.Text11.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyLanReEst.NameE  like '%" & Me.Text11.text & "%'"
        End If
    End If
    

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblBuyLanReEst.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid4
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("No_planned")) = IIf(IsNull(rs("No_planned").value), "", rs("No_planned").value)
                .TextMatrix(i, .ColIndex("Area")) = IIf(IsNull(rs("Area").value), "", rs("Area").value)
                .TextMatrix(i, .ColIndex("MeterPrice")) = IIf(IsNull(rs("MeterPrice").value), "", rs("MeterPrice").value)
                .TextMatrix(i, .ColIndex("TitledeedNo")) = IIf(IsNull(rs("TitledeedNo").value), "", rs("TitledeedNo").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                If Not (IsNull(rs("PayType").value)) Then
                If rs("PayType").value = 0 Then
                .TextMatrix(i, .ColIndex("PayType")) = "‰ﬁœÌ"
                ElseIf rs("PayType").value = 1 Then
                .TextMatrix(i, .ColIndex("PayType")) = "«Ã·"
                End If
                End If
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                    If Not (IsNull(rs("PayType").value)) Then
                If rs("PayType").value = 0 Then
                .TextMatrix(i, .ColIndex("PayType")) = "Cash"
                ElseIf rs("PayType").value = 1 Then
                .TextMatrix(i, .ColIndex("PayType")) = "Credit"
                End If
                End If
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataExpensesLand()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  sql = "SELECT     dbo.TblExpensesInvesment.ID, dbo.TblExpensesInvesment.RecordDate, dbo.TblExpensesInvesment.CurrValue, dbo.TblExpensesInvesment.DevlopValue, "
  sql = sql & "                    dbo.TblExpensesInvesment.AfterDevlopValue, dbo.TblExpensesInvesment.ShareValue, dbo.TblExpensesInvesment.SharNo, dbo.TblExpensesInvesment.Remarks,"
  sql = sql & "                    dbo.TblExpensesInvesment.Total, dbo.TblExpensesInvesment.ShareValueNew, dbo.TblExpensesInvesmentDet.StartDate, dbo.TblExpensesInvesmentDet.EndDate,"
  sql = sql & "                    dbo.TblExpensesInvesmentDet.FromArea, dbo.TblExpensesInvesmentDet.ToArea, dbo.TblExpensesInvesmentDet.Valu,"
  sql = sql & "                    dbo.TblExpensesInvesmentDet.Remarks AS RemarksDet, dbo.TblExpensesInvesment.BranchID, dbo.TblBranchesData.branch_name,"
  sql = sql & "                    dbo.TblBranchesData.branch_namee, dbo.TblExpensesInvesment.InvesID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblExpensesInvesment.LandID,"
  sql = sql & "                    dbo.TblBuyLanReEst.Name AS LandName, dbo.TblBuyLanReEst.NameE AS LandNameE, dbo.TblExpensesInvesmentDet.Cus_ID, dbo.TblCustemers.CusName,"
  sql = sql & "                     dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblExpensesInvesmentDet.DevlopID, dbo.TblInvestmentsGroup.Name AS InvName,"
  sql = sql & "                     dbo.TblInvestmentsGroup.NameE AS InvNameE ,dbo.TblExpensesInvesmentDet.TypTrns"
  sql = sql & " FROM         dbo.TblExpensesInvesmentDet LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblInvestmentsGroup ON dbo.TblExpensesInvesmentDet.DevlopID = dbo.TblInvestmentsGroup.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblCustemers ON dbo.TblExpensesInvesmentDet.Cus_ID = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblExpensesInvesment LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblBuyLanReEst ON dbo.TblExpensesInvesment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.Tblinvestment ON dbo.TblExpensesInvesment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblBranchesData ON dbo.TblExpensesInvesment.BranchID = dbo.TblBranchesData.branch_id ON"
  sql = sql & "                    dbo.TblExpensesInvesmentDet.ExpInvID = dbo.TblExpensesInvesment.ID"
       BolBegine = False
       StrWhere = ""
  
    
   ''''''''''''''''''''''''''''''''''''
    
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblExpensesInvesment.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblExpensesInvesment.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblExpensesInvesment.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblExpensesInvesment.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblExpensesInvesment.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblExpensesInvesment.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblExpensesInvesment.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblExpensesInvesment.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcbInvise.text <> "" And (val(DcbInvise.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesment.InvesID =" & val(Me.DcbInvise.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesment.InvesID =" & val(Me.DcbInvise.BoundText) & ""
       End If
    End If
    If Me.DcbLand.text <> "" And (val(DcbLand.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesment.LandID =" & val(Me.DcbLand.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesment.LandID =" & val(Me.DcbLand.BoundText) & ""
       End If
     End If
     If Me.DcbCus.text <> "" And (val(DcbCus.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesmentDet.Cus_ID =" & val(Me.DcbCus.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesmentDet.Cus_ID =" & val(Me.DcbCus.BoundText) & ""
       End If
     End If
     
         If Me.DbDevlop.text <> "" And (val(DbDevlop.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesmentDet.DevlopID =" & val(Me.DbDevlop.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesmentDet.DevlopID =" & val(Me.DbDevlop.BoundText) & ""
       End If
     End If

       If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesmentDet.TypTrns =1"
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesmentDet.TypTrns =1"
       End If
     
    

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblExpensesInvesment.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid5
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InvesID")) = IIf(IsNull(rs("InvesID").value), "", rs("InvesID").value)
                .TextMatrix(i, .ColIndex("LandID")) = IIf(IsNull(rs("LandID").value), "", rs("LandID").value)
                .TextMatrix(i, .ColIndex("DevlopID")) = IIf(IsNull(rs("DevlopID").value), "", rs("DevlopID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("LandName")) = IIf(IsNull(rs("LandName").value), "", rs("LandName").value)
                .TextMatrix(i, .ColIndex("InvName")) = IIf(IsNull(rs("InvName").value), "", rs("InvName").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("LandName")) = IIf(IsNull(rs("LandNameE").value), "", rs("LandNameE").value)
                .TextMatrix(i, .ColIndex("InvName")) = IIf(IsNull(rs("InvNameE").value), "", rs("InvNameE").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataExpensesLand2()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  sql = "SELECT        dbo.TblExpensesInvesmentDet.StartDate, dbo.TblExpensesInvesmentDet.EndDate, dbo.TblExpensesInvesmentDet.FromArea, dbo.TblExpensesInvesmentDet.ToArea, dbo.TblExpensesInvesmentDet.Valu, "
  sql = sql & "                        dbo.TblExpensesInvesmentDet.Remarks AS RemarksDet, dbo.TblExpensesInvesmentDet.Cus_ID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
  sql = sql & "                        dbo.TblExpensesInvesmentDet.DevlopID, dbo.TblInvestmentsGroup.Name AS InvName, dbo.TblInvestmentsGroup.NameE AS InvNameE, dbo.TblReturnExpensInves.ID, dbo.TblReturnExpensInves.RecordDate,"
  sql = sql & "                        dbo.TblReturnExpensInves.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblReturnExpensInves.Remarks, dbo.TblReturnExpensInves.InvesID,"
  sql = sql & "                        dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblReturnExpensInves.LandID, dbo.TblBuyLanReEst.Name AS LandName, dbo.TblBuyLanReEst.NameE AS LandNameE,"
  sql = sql & "                        dbo.TblExpensesInvesmentDet.TypTrns"
  sql = sql & "  FROM            dbo.TblCustemers RIGHT OUTER JOIN"
  sql = sql & "                        dbo.TblInvestmentsGroup RIGHT OUTER JOIN"
  sql = sql & "                        dbo.Tblinvestment RIGHT OUTER JOIN"
  sql = sql & "                        dbo.TblBuyLanReEst INNER JOIN"
  sql = sql & "                        dbo.TblReturnExpensInves ON dbo.TblBuyLanReEst.ID = dbo.TblReturnExpensInves.LandID LEFT OUTER JOIN"
  sql = sql & "                        dbo.TblExpensesInvesmentDet ON dbo.TblReturnExpensInves.ID = dbo.TblExpensesInvesmentDet.ExpInvID ON dbo.Tblinvestment.ID = dbo.TblReturnExpensInves.InvesID LEFT OUTER JOIN"
  sql = sql & "                        dbo.TblBranchesData ON dbo.TblReturnExpensInves.ID = dbo.TblBranchesData.branch_id ON dbo.TblInvestmentsGroup.ID = dbo.TblExpensesInvesmentDet.DevlopID ON"
  sql = sql & "                        dbo.TblCustemers.CusID = dbo.TblExpensesInvesmentDet.Cus_ID"
       BolBegine = False
       StrWhere = ""
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblReturnExpensInves.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblReturnExpensInves.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblReturnExpensInves.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblReturnExpensInves.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblReturnExpensInves.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblReturnExpensInves.RecordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblReturnExpensInves.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblReturnExpensInves.RecordDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcbInvise.text <> "" And (val(DcbInvise.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblReturnExpensInves.InvesID =" & val(Me.DcbInvise.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblReturnExpensInves.InvesID =" & val(Me.DcbInvise.BoundText) & ""
       End If
    End If
    If Me.DcbLand.text <> "" And (val(DcbLand.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblReturnExpensInves.LandID =" & val(Me.DcbLand.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblReturnExpensInves.LandID =" & val(Me.DcbLand.BoundText) & ""
       End If
     End If
     If Me.DcbCus.text <> "" And (val(DcbCus.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesmentDet.Cus_ID =" & val(Me.DcbCus.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesmentDet.Cus_ID =" & val(Me.DcbCus.BoundText) & ""
       End If
     End If
     
         If Me.DbDevlop.text <> "" And (val(DbDevlop.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesmentDet.DevlopID =" & val(Me.DbDevlop.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesmentDet.DevlopID =" & val(Me.DbDevlop.BoundText) & ""
       End If
     End If

       If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblExpensesInvesmentDet.TypTrns =-1"
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblExpensesInvesmentDet.TypTrns =-1"
       End If
     
    

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblReturnExpensInves.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid5
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InvesID")) = IIf(IsNull(rs("InvesID").value), "", rs("InvesID").value)
                .TextMatrix(i, .ColIndex("LandID")) = IIf(IsNull(rs("LandID").value), "", rs("LandID").value)
                .TextMatrix(i, .ColIndex("DevlopID")) = IIf(IsNull(rs("DevlopID").value), "", rs("DevlopID").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("LandName")) = IIf(IsNull(rs("LandName").value), "", rs("LandName").value)
                .TextMatrix(i, .ColIndex("InvName")) = IIf(IsNull(rs("InvName").value), "", rs("InvName").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("LandName")) = IIf(IsNull(rs("LandNameE").value), "", rs("LandNameE").value)
                .TextMatrix(i, .ColIndex("InvName")) = IIf(IsNull(rs("InvNameE").value), "", rs("InvNameE").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataDivideLand()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  sql = "SELECT     dbo.TblDivInvesment.ID, dbo.TblDivInvesment.RecordDate, dbo.TblDivInvesment.CurrValue, dbo.TblDivInvesment.DevlopValue, "
  sql = sql & "                   dbo.TblDivInvesment.AfterDevlopValue, dbo.TblDivInvesment.ShareValue, dbo.TblDivInvesment.SharNo, dbo.TblDivInvesment.Remarks, dbo.TblDivInvesment.Total,"
  sql = sql & "                     dbo.TblDivInvesment.AlwArea, dbo.TblDivInvesment.AlwAreaAfter, dbo.TblDivInvesment.BranchID, dbo.TblBranchesData.branch_name,"
  sql = sql & "                     dbo.TblBranchesData.branch_namee, dbo.TblDivInvesment.InvesID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblDivInvesment.LandID,"
  sql = sql & "                     dbo.TblBuyLanReEst.Name AS LandName, dbo.TblBuyLanReEst.NameE AS LandNameE, dbo.TblDivInvesmentDet.TypeDivi, dbo.TblSpreading.Name AS DivName,"
  sql = sql & "                     dbo.TblSpreading.NameE AS DivNameE, dbo.TblDivInvesmentDet.EffectID, dbo.TblDivInvesmentDet.Area, dbo.TblDivInvesmentDet.TotalArea,"
  sql = sql & "                     dbo.TblDivInvesmentDet.NewArea, dbo.TblDivInvesmentDet.PartNo, dbo.TblDivInvesmentDet.BlokNo, dbo.TblDivInvesmentDet.Nourth,"
  sql = sql & "                     dbo.TblDivInvesmentDet.South , dbo.TblDivInvesmentDet.East, dbo.TblDivInvesmentDet.West"
  sql = sql & "   FROM         dbo.TblSpreading RIGHT OUTER JOIN"
  sql = sql & "                     dbo.TblDivInvesmentDet ON dbo.TblSpreading.ID = dbo.TblDivInvesmentDet.TypeDivi RIGHT OUTER JOIN"
  sql = sql & "                     dbo.TblDivInvesment ON dbo.TblDivInvesmentDet.DivInvID = dbo.TblDivInvesment.ID LEFT OUTER JOIN"
  sql = sql & "                     dbo.TblBuyLanReEst ON dbo.TblDivInvesment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
  sql = sql & "                     dbo.Tblinvestment ON dbo.TblDivInvesment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  sql = sql & "                     dbo.TblBranchesData ON dbo.TblDivInvesment.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = ""
   ''''''''''''''''''''''''''''''''''''
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblDivInvesment.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDivInvesment.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblDivInvesment.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblDivInvesment.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDivInvesment.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDivInvesment.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDivInvesment.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblDivInvesment.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcbInvise.text <> "" And (val(DcbInvise.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblDivInvesment.InvesID =" & val(Me.DcbInvise.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblDivInvesment.InvesID =" & val(Me.DcbInvise.BoundText) & ""
       End If
    End If
    If Me.DcbLand.text <> "" And (val(DcbLand.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblDivInvesment.LandID =" & val(Me.DcbLand.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblDivInvesment.LandID =" & val(Me.DcbLand.BoundText) & ""
       End If
     End If
   If Me.DcbDiv.text <> "" And (val(DcbDiv.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblDivInvesmentDet.TypeDivi =" & val(Me.DcbDiv.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblDivInvesmentDet.TypeDivi =" & val(Me.DcbDiv.BoundText) & ""
       End If
     End If
     
     If Me.TxtBlockNo.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDivInvesmentDet.BlokNo like '%" & Me.TxtBlockNo.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDivInvesmentDet.BlokNo  like '%" & Me.TxtBlockNo.text & "%'"
        End If
    End If
     If Me.TxtPartNo.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblDivInvesmentDet.PartNo like '%" & Me.TxtPartNo.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblDivInvesmentDet.PartNo  like '%" & Me.TxtPartNo.text & "%'"
        End If
    End If

     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblDivInvesment.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid7
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InvesID")) = IIf(IsNull(rs("InvesID").value), "", rs("InvesID").value)
                .TextMatrix(i, .ColIndex("LandID")) = IIf(IsNull(rs("LandID").value), "", rs("LandID").value)
                .TextMatrix(i, .ColIndex("BlokNo")) = IIf(IsNull(rs("BlokNo").value), "", rs("BlokNo").value)
                .TextMatrix(i, .ColIndex("TypeDivi")) = IIf(IsNull(rs("TypeDivi").value), "", rs("TypeDivi").value)
                .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(rs("PartNo").value), "", rs("PartNo").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("LandName")) = IIf(IsNull(rs("LandName").value), "", rs("LandName").value)
                '.TextMatrix(i, .ColIndex("InvName")) = IIf(IsNull(rs("InvName").value), "", rs("InvName").value)
                .TextMatrix(i, .ColIndex("DivName")) = IIf(IsNull(rs("DivName").value), "", rs("DivName").value)
                Else
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("LandName")) = IIf(IsNull(rs("LandNameE").value), "", rs("LandNameE").value)
                '.TextMatrix(i, .ColIndex("DivName")) = IIf(IsNull(rs("DivNameE").value), "", rs("DivNameE").value)
                .TextMatrix(i, .ColIndex("DivName")) = IIf(IsNull(rs("DivNameE").value), "", rs("DivNameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataBuyBillInvestment()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  sql = "SELECT     dbo.TblBuyBilllInvestment.ID, dbo.TblBuyBilllInvestment.RecordDate, dbo.TblBuyBilllInvestment.Payment, dbo.TblBuyBilllInvestment.RecordNo, "
  sql = sql & "                    dbo.TblBuyBilllInvestment.CusID, dbo.TblBuyBilllInvestment.SharNo, dbo.TblBuyBilllInvestment.Remarks, dbo.TblBuyBilllInvestment.SharValue,"
  sql = sql & "                    dbo.TblBuyBilllInvestment.Cus_ID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblBuyBilllInvestment.BranchID,"
  sql = sql & "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblBuyBilllInvestment.UserID, dbo.TblBuyBilllInvestment.SellerID,"
  sql = sql & "                    TblCustemers_1.CusName AS SalerCusName, TblCustemers_1.CusNamee AS SalerCusNameE, TblCustemers_1.Fullcode AS SalFullcode,"
  sql = sql & "                    dbo.TblBuyBilllInvestment.Cus_Type, dbo.TblInvestorType.Name, dbo.TblInvestorType.NameE, dbo.TblInvestorType.Code, dbo.TblBuyBilllInvestment.InvesID,"
  sql = sql & "                    dbo.Tblinvestment.Name AS InvName, dbo.Tblinvestment.NameE AS InvNameE, dbo.TblBuyBilllInvestmentDet.Remarks AS DetRemarks,"
  sql = sql & "                    dbo.TblBuyBilllInvestmentDet.SharNo AS DetSharNo, dbo.TblBuyBilllInvestmentDet.SharValue AS DetSharValue, dbo.TblBuyBilllInvestmentDet.Total,"
  sql = sql & "                    dbo.TblBuyBilllInvestmentDet.BeforTotal, dbo.TblBuyBilllInvestmentDet.Profit, dbo.TblBuyBilllInvestmentDet.SharValueBefor,"
  sql = sql & "                    dbo.TblBuyBilllInvestmentDet.InvesID AS DetInvesID, Tblinvestment_1.Name AS DetInvName, Tblinvestment_1.NameE AS DetInvNameE"
  sql = sql & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblBranchesData RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblBuyBilllInvestment LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblBuyBilllInvestmentDet ON dbo.TblBuyBilllInvestment.ID = dbo.TblBuyBilllInvestmentDet.BuyBilInvsID LEFT OUTER JOIN"
  sql = sql & "                    dbo.Tblinvestment Tblinvestment_1 ON dbo.TblBuyBilllInvestmentDet.InvesID = Tblinvestment_1.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.Tblinvestment ON dbo.TblBuyBilllInvestment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblInvestorType ON dbo.TblBuyBilllInvestment.Cus_Type = dbo.TblInvestorType.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblCustemers TblCustemers_1 ON dbo.TblBuyBilllInvestment.SellerID = TblCustemers_1.CusID ON"
  sql = sql & "                    dbo.TblBranchesData.branch_id = dbo.TblBuyBilllInvestment.BranchID ON dbo.TblCustemers.CusID = dbo.TblBuyBilllInvestment.Cus_ID"
       BolBegine = False
       StrWhere = ""
   ''''''''''''''''''''''''''''''''''''
         If val(Me.TxtSharNo.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblBuyBilllInvestmentDet.SharNo >=" & val(Me.TxtSharNo.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyBilllInvestmentDet.SharNo >=" & val(Me.TxtSharNo.text) & ""
        End If
    End If
    If val(Me.TxtSharNoto.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblBuyBilllInvestmentDet.SharNo <=" & val(Me.TxtSharNoto.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblBuyBilllInvestmentDet.SharNo <=" & val(Me.TxtSharNoto.text) & ""
       End If
    End If
   ''/////////////////////////
     If val(Me.TxtSharValue.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblBuyBilllInvestmentDet.SharValue >=" & val(Me.TxtSharValue.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyBilllInvestmentDet.SharValue >=" & val(Me.TxtSharValue.text) & ""
        End If
    End If
    If val(Me.TxtSharValueTo.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblBuyBilllInvestmentDet.SharValue <=" & val(Me.TxtSharValueTo.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblBuyBilllInvestmentDet.SharValue <=" & val(Me.TxtSharValueTo.text) & ""
       End If
    End If
   '''////////
    
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblBuyBilllInvestment.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyBilllInvestment.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblBuyBilllInvestment.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblBuyBilllInvestment.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblBuyBilllInvestment.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblBuyBilllInvestment.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblBuyBilllInvestment.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblBuyBilllInvestment.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcbInves.text <> "" And (val(DcbInves.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblBuyBilllInvestmentDet.InvesID =" & val(Me.DcbInves.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblBuyBilllInvestmentDet.InvesID =" & val(Me.DcbInves.BoundText) & ""
       End If
    End If
    If Me.DcbCusto.text <> "" And (val(DcbCusto.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblBuyBilllInvestment.Cus_ID =" & val(Me.DcbCusto.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblBuyBilllInvestment.Cus_ID =" & val(Me.DcbCusto.BoundText) & ""
       End If
     End If
   If Me.DcbSales.text <> "" And (val(DcbSales.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  dbo.TblBuyBilllInvestment.SellerID =" & val(Me.DcbSales.BoundText) & ""
        Else:
          BolBegine = True
          StrWhere = " Where dbo.TblBuyBilllInvestment.SellerID =" & val(Me.DcbSales.BoundText) & ""
       End If
     End If
     
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblBuyBilllInvestment.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid8
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordNo").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordNo").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InvesID")) = IIf(IsNull(rs("DetInvesID").value), "", rs("DetInvesID").value)
                .TextMatrix(i, .ColIndex("DetSharValue")) = IIf(IsNull(rs("DetSharValue").value), "", rs("DetSharValue").value)
                .TextMatrix(i, .ColIndex("DetSharNo")) = IIf(IsNull(rs("DetSharNo").value), "", rs("DetSharNo").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("SalerCusName")) = IIf(IsNull(rs("SalerCusName").value), "", rs("SalerCusName").value)
                .TextMatrix(i, .ColIndex("DetInvName")) = IIf(IsNull(rs("DetInvName").value), "", rs("DetInvName").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                .TextMatrix(i, .ColIndex("SalerCusName")) = IIf(IsNull(rs("SalerCusNameE").value), "", rs("SalerCusNameE").value)
                .TextMatrix(i, .ColIndex("DetInvName")) = IIf(IsNull(rs("DetInvNameE").value), "", rs("DetInvNameE").value)
               .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataSalesBill()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  sql = "SELECT     dbo.TblSaleBilllInvestment.ID, dbo.TblSaleBilllInvestment.RecordDate, dbo.TblSaleBilllInvestment.UserID, dbo.TblSaleBilllInvestment.SellerType, "
  sql = sql & "                    dbo.TblSaleBilllInvestment.commission, dbo.TblSaleBilllInvestment.DesLocation, dbo.TblSaleBilllInvestment.Remarks, dbo.TblSaleBilllInvestment.PropertyDeed,"
  sql = sql & "                    dbo.TblSaleBilllInvestment.NorthlengthStr, dbo.TblSaleBilllInvestment.SouthlengthStr, dbo.TblSaleBilllInvestment.EastlengthStr,"
  sql = sql & "                    dbo.TblSaleBilllInvestment.WestlengthStr, dbo.TblSaleBilllInvestment.Northlength, dbo.TblSaleBilllInvestment.Southlength, dbo.TblSaleBilllInvestment.Eastlength,"
  sql = sql & "                    dbo.TblSaleBilllInvestment.Westlength, dbo.TblSaleBilllInvestment.Payment, dbo.TblSaleBilllInvestment.RecordNo, dbo.TblSaleBilllInvestment.CusID,"
  sql = sql & "                    dbo.TblSaleBilllInvestment.PaymentNo, dbo.TblSaleBilllInvestment.Period, dbo.TblSaleBilllInvestment.PeriodType, dbo.TblSaleBilllInvestment.RemarkPay,"
  sql = sql & "                    dbo.TblSaleBilllInvestment.FristDate, dbo.TblSaleBilllInvestment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
  sql = sql & "                    dbo.TblSaleBilllInvestment.SellerID, dbo.TblCustemers.CusName AS SelCusName, dbo.TblCustemers.CusNamee AS SelCusNameE,"
  sql = sql & "                    dbo.TblCustemers.Fullcode AS SelFullcode, dbo.TblEmployee.Emp_Name AS SelEmp_Name, dbo.TblEmployee.Fullcode AS SelEmpFullcode,"
  sql = sql & "                    dbo.TblEmployee.Emp_Namee AS SelEmp_NameE, dbo.TblSaleBilllInvestment.LandID, dbo.TblBuyLanReEst.Name, dbo.TblBuyLanReEst.NameE,"
  sql = sql & "                    dbo.TblSaleBilllInvestment.Cus_ID, TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode, dbo.TblSaleBilllInvestment.Cus_Tpe,"
  sql = sql & "                    dbo.TblInvestorType.Name AS TypName, dbo.TblInvestorType.NameE AS TypNameE, dbo.TblSaleBilllInvestmentDet.FristDate AS DetFristDate,"
  sql = sql & "                    dbo.TblSaleBilllInvestmentDet.Profit, dbo.TblSaleBilllInvestmentDet.TotalCost, dbo.TblSaleBilllInvestmentDet.MeterValue, dbo.TblSaleBilllInvestmentDet.Payed,"
  sql = sql & "                    dbo.TblSaleBilllInvestmentDet.Remarks AS DetRemarks, dbo.TblSaleBilllInvestmentDet.Net, dbo.TblSaleBilllInvestmentDet.TypeDis,"
  sql = sql & "                    dbo.TblSaleBilllInvestmentDet.DisValue, dbo.TblSaleBilllInvestmentDet.Total, dbo.TblSaleBilllInvestmentDet.Valu, dbo.TblSaleBilllInvestmentDet.Area,"
  sql = sql & "                    dbo.TblSaleBilllInvestmentDet.TypeTrns, dbo.TblSaleBilllInvestmentDet.InvesID, dbo.Tblinvestment.Name AS InvestName,"
  sql = sql & "                    dbo.Tblinvestment.NameE AS InvestNameE, dbo.TblSaleBilllInvestmentDet.DivID, dbo.TblSaleBilllInvestmentDet.PartID, dbo.TblDivInvesmentDet.PartNo"
  sql = sql & "    FROM         dbo.TblDivInvesmentDet RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblSaleBilllInvestmentDet ON dbo.TblDivInvesmentDet.ID = dbo.TblSaleBilllInvestmentDet.PartID LEFT OUTER JOIN"
  sql = sql & "                    dbo.Tblinvestment ON dbo.TblSaleBilllInvestmentDet.InvesID = dbo.Tblinvestment.ID RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblSaleBilllInvestment ON dbo.TblSaleBilllInvestmentDet.SBINVID = dbo.TblSaleBilllInvestment.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblInvestorType ON dbo.TblSaleBilllInvestment.Cus_Tpe = dbo.TblInvestorType.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblCustemers TblCustemers_1 ON dbo.TblSaleBilllInvestment.Cus_ID = TblCustemers_1.CusID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblBuyLanReEst ON dbo.TblSaleBilllInvestment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblEmployee ON dbo.TblSaleBilllInvestment.SellerID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblCustemers ON dbo.TblSaleBilllInvestment.SellerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblBranchesData ON dbo.TblSaleBilllInvestment.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = " WHERE     (dbo.TblSaleBilllInvestmentDet.TypeTrns = 0)"
   ''''''''''''''''''''''''''''''''''''
         If val(Me.Text19.text) <> 0 Then
            StrWhere = StrWhere & "AND dbo.TblSaleBilllInvestmentDet.Valu >=" & val(Me.Text19.text) & ""
    End If
    If val(Me.Text21.text) <> 0 Then
         StrWhere = StrWhere & "AND  dbo.TblSaleBilllInvestmentDet.Valu <=" & val(Me.Text21.text) & ""
    End If
    If val(Me.TxtIDFrom.text) <> 0 Then
            StrWhere = StrWhere & "AND  dbo.TblSaleBilllInvestment.ID >=" & val(Me.TxtIDFrom.text) & ""
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
          StrWhere = StrWhere & " AND dbo.TblSaleBilllInvestment.ID <=" & val(Me.TxtIDTO.text) & ""
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
            StrWhere = StrWhere & " AND dbo.TblSaleBilllInvestment.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND dbo.TblSaleBilllInvestment.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If Me.DcbTypeSales.text <> "" And (val(DcbTypeSales.ListIndex) <> -1) Then
           StrWhere = StrWhere & " AND  dbo.TblSaleBilllInvestment.SellerType =" & val(Me.DcbTypeSales.ListIndex) & ""
    End If
    If Me.DcbSeler.text <> "" And (val(DcbSeler.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblSaleBilllInvestment.SellerID =" & val(Me.DcbSeler.BoundText) & ""
    End If
     If Me.DcbInves3.text <> "" And (val(DcbInves3.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblSaleBilllInvestmentDet.InvesID =" & val(Me.DcbInves3.BoundText) & ""
    End If
    If Me.DcCustomerType.text <> "" And (val(DcCustomerType.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblSaleBilllInvestment.Cus_Tpe =" & val(Me.DcCustomerType.BoundText) & ""
    End If
      If Me.DcbCust2.text <> "" And (val(DcbCust2.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblSaleBilllInvestment.Cus_ID =" & val(Me.DcbCust2.BoundText) & ""
    End If
          If Me.DcbLand1.text <> "" And (val(DcbLand1.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblSaleBilllInvestment.LandID =" & val(Me.DcbLand1.BoundText) & ""
    End If
      If Me.TxtPart11.text <> "" Then
            StrWhere = StrWhere & "AND  dbo.TblDivInvesmentDet.PartNo like '%" & Me.TxtPart11.text & "%'"
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblSaleBilllInvestment.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid9
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(rs("PartNo").value), "", rs("PartNo").value)
                .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs("Valu").value), "", rs("Valu").value)
                DcbTypeSales2.ListIndex = IIf(IsNull(rs("SellerType").value), -1, rs("SellerType").value)
                .TextMatrix(i, .ColIndex("SellerType")) = DcbTypeSales2.text
               If SystemOptions.UserInterface = ArabicInterface Then
               If DcbTypeSales2.ListIndex = 0 Then
               .TextMatrix(i, .ColIndex("SelCusName")) = IIf(IsNull(rs("SelEmp_Name").value), "", rs("SelEmp_Name").value)
               Else
               .TextMatrix(i, .ColIndex("SelCusName")) = IIf(IsNull(rs("SelCusName").value), "", rs("SelCusName").value)
               End If
               .TextMatrix(i, .ColIndex("InvestName")) = IIf(IsNull(rs("InvestName").value), "", rs("InvestName").value)
                .TextMatrix(i, .ColIndex("TypName")) = IIf(IsNull(rs("TypName").value), "", rs("TypName").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                If DcbTypeSales2.ListIndex = 0 Then
               .TextMatrix(i, .ColIndex("SelCusName")) = IIf(IsNull(rs("SelEmp_NameE").value), "", rs("SelEmp_NameE").value)
               Else
               .TextMatrix(i, .ColIndex("SelCusName")) = IIf(IsNull(rs("SelCusNameE").value), "", rs("SelCusNameE").value)
               End If
                .TextMatrix(i, .ColIndex("InvestName")) = IIf(IsNull(rs("InvestNameE").value), "", rs("InvestNameE").value)
                .TextMatrix(i, .ColIndex("TypName")) = IIf(IsNull(rs("TypNameE").value), "", rs("TypNameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataInvestProfit()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  sql = "SELECT     dbo.TblInvestProfitDistri.ID, dbo.TblInvestProfitDistri.RecordDate, dbo.TblInvestProfitDistri.Remarks, dbo.TblInvestProfitDistri.Comm, "
  sql = sql & "                    dbo.TblInvestProfitDistri.InvestValue, dbo.TblInvestProfitDistri.SalValue, dbo.TblInvestProfitDistri.PorfetValue, dbo.TblInvestProfitDistri.SharNo,"
  sql = sql & "                    dbo.TblInvestProfitDistri.TotalShare, dbo.TblInvestProfitDistri.UserID, dbo.TblInvestProfitDistri.BranchID, dbo.TblBranchesData.branch_name,"
  sql = sql & "                    dbo.TblBranchesData.branch_namee, dbo.TblInvestProfitDistri.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
  sql = sql & "                    dbo.TblInvestProfitDistri.InvesID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblInvestProfitDistri.TypeShere, dbo.TblInvestProfitDistri.ShareID,"
  sql = sql & "                    dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblInvestProfitDistriDet.Remarks AS DetRemarks,"
  sql = sql & "                    dbo.TblInvestProfitDistriDet.SharNo AS DetSharNo, dbo.TblInvestProfitDistriDet.Profit, dbo.TblInvestProfitDistriDet.TypeTrans, dbo.TblInvestProfitDistriDet.SBINVID,"
  sql = sql & "                    dbo.TblInvestProfitDistriDet.BilValue, dbo.TblInvestProfitDistriDet.RecordDate AS DetRecordDate, dbo.TblInvestProfitDistriDet.ShareID AS DetShareID,"
  sql = sql & "                    TblCustemers_1.CusName AS DetCusName, TblCustemers_1.CusNamee AS DetCusNamee, TblCustemers_1.Fullcode AS DetFullcode"
  sql = sql & "  FROM         dbo.TblCustemers TblCustemers_1 RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblInvestProfitDistriDet ON TblCustemers_1.CusID = dbo.TblInvestProfitDistriDet.ShareID RIGHT OUTER JOIN"
  sql = sql & "                    dbo.TblInvestProfitDistri ON dbo.TblInvestProfitDistriDet.InvProID = dbo.TblInvestProfitDistri.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblCustemers ON dbo.TblInvestProfitDistri.ShareID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  sql = sql & "                    dbo.Tblinvestment ON dbo.TblInvestProfitDistri.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblEmployee ON dbo.TblInvestProfitDistri.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblBranchesData ON dbo.TblInvestProfitDistri.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = " WHERE     (dbo.TblInvestProfitDistriDet.TypeTrans = 0)"
   ''''''''''''''''''''''''''''''''''''
     If val(Me.Text23.text) <> 0 Then
            StrWhere = StrWhere & "AND dbo.TblInvestProfitDistri.InvestValue >=" & val(Me.Text23.text) & ""
    End If
    If val(Me.Text20.text) <> 0 Then
         StrWhere = StrWhere & "AND  dbo.TblInvestProfitDistri.InvestValue <=" & val(Me.Text20.text) & ""
    End If
    ''////////////////////
      If val(Me.Text28.text) <> 0 Then
            StrWhere = StrWhere & "AND dbo.TblInvestProfitDistri.SalValue >=" & val(Me.Text28.text) & ""
    End If
    If val(Me.Text27.text) <> 0 Then
         StrWhere = StrWhere & "AND  dbo.TblInvestProfitDistri.SalValue <=" & val(Me.Text27.text) & ""
    End If
    '''
        ''////////////////////
      If val(Me.Text25.text) <> 0 Then
            StrWhere = StrWhere & "AND dbo.TblInvestProfitDistri.PorfetValue >=" & val(Me.Text25.text) & ""
    End If
    If val(Me.Text22.text) <> 0 Then
         StrWhere = StrWhere & "AND  dbo.TblInvestProfitDistri.PorfetValue <=" & val(Me.Text22.text) & ""
    End If
    '''
    If val(Me.TxtIDFrom.text) <> 0 Then
            StrWhere = StrWhere & "AND  dbo.TblInvestProfitDistri.ID >=" & val(Me.TxtIDFrom.text) & ""
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
          StrWhere = StrWhere & " AND dbo.TblInvestProfitDistri.ID <=" & val(Me.TxtIDTO.text) & ""
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
            StrWhere = StrWhere & " AND dbo.TblInvestProfitDistri.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND dbo.TblInvestProfitDistri.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If Me.DcbEmp.text <> "" And (val(DcbEmp.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblInvestProfitDistri.EmpID =" & val(Me.DcbEmp.BoundText) & ""
    End If
    If Me.DcbInves4.text <> "" And (val(DcbInves4.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblInvestProfitDistri.InvesID =" & val(Me.DcbInves4.BoundText) & ""
    End If
     If Me.DcbSharer.text <> "" And (val(DcbSharer.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblInvestProfitDistriDet.ShareID =" & val(Me.DcbSharer.BoundText) & ""
    End If
  
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblInvestProfitDistri.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid10
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InvestValue")) = IIf(IsNull(rs("InvestValue").value), "", rs("InvestValue").value)
                .TextMatrix(i, .ColIndex("SalValue")) = IIf(IsNull(rs("SalValue").value), "", rs("SalValue").value)
                .TextMatrix(i, .ColIndex("PorfetValue")) = IIf(IsNull(rs("PorfetValue").value), -1, rs("PorfetValue").value)
               
               If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("DetCusName")) = IIf(IsNull(rs("DetCusName").value), "", rs("DetCusName").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("DetCusName")) = IIf(IsNull(rs("DetCusNamee").value), "", rs("DetCusNamee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataInvestLiquidation()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  sql = "SELECT     dbo.TblInvestliquidation.ID, dbo.TblInvestliquidation.RecordDate, dbo.TblInvestliquidation.LiqDate, dbo.TblInvestliquidation.BranchID, "
  sql = sql & "                     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblInvestliquidation.Remarks, dbo.TblInvestliquidation.TotalSalLand,"
  sql = sql & "                    dbo.TblInvestliquidation.TotalExpens, dbo.TblInvestliquidation.TotalCost, dbo.TblInvestliquidation.TotalShare, dbo.TblInvestliquidation.TotalDiff,"
  sql = sql & "                    dbo.TblInvestliquidation.WritSalLand, dbo.TblInvestliquidation.WritExpens, dbo.TblInvestliquidation.WritCost, dbo.TblInvestliquidation.WritShare,"
  sql = sql & "                    dbo.TblInvestliquidation.WritDiff, dbo.TblInvestliquidation.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
  sql = sql & "                    dbo.TblInvestliquidation.InvesID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.Tblinvestment.ID AS InvesmentID"
  sql = sql & "  FROM         dbo.TblInvestliquidation LEFT OUTER JOIN"
  sql = sql & "                    dbo.Tblinvestment ON dbo.TblInvestliquidation.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  sql = sql & "                    dbo.TblEmployee ON dbo.TblInvestliquidation.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  sql = sql & "                   dbo.TblBranchesData ON dbo.TblInvestliquidation.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = " WHERE     (1=1)"
   ''''''''''''''''''''''''''''''''''''

    If val(Me.TxtIDFrom.text) <> 0 Then
            StrWhere = StrWhere & "AND  dbo.TblInvestliquidation.ID >=" & val(Me.TxtIDFrom.text) & ""
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
          StrWhere = StrWhere & " AND dbo.TblInvestliquidation.ID <=" & val(Me.TxtIDTO.text) & ""
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
            StrWhere = StrWhere & " AND dbo.TblInvestliquidation.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND dbo.TblInvestliquidation.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If
       If Not IsNull(Me.FromLiqDate.value) Then
            StrWhere = StrWhere & " AND dbo.TblInvestliquidation.LiqDate >=" & SQLDate(Me.FromLiqDate.value, True) & ""
    End If
    If Not IsNull(Me.ToLiqDate.value) Then
            StrWhere = StrWhere & " AND dbo.TblInvestliquidation.LiqDate <=" & SQLDate(Me.ToLiqDate.value, True) & ""
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       If Me.DcbEmployee.text <> "" And (val(DcbEmployee.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblInvestliquidation.EmpID =" & val(Me.DcbEmployee.BoundText) & ""
    End If
    If Me.DcbInvest.text <> "" And (val(DcbInvest.BoundText) <> 0) Then
           StrWhere = StrWhere & " AND  dbo.TblInvestliquidation.InvesID =" & val(Me.DcbInvest.BoundText) & ""
    End If
       
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblInvestliquidation.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
      
        End If
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid11
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                  If Not (IsNull(rs("LiqDate").value)) Then
                  .TextMatrix(i, .ColIndex("LiqDate")) = Format(rs("LiqDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("InvesmentID")) = IIf(IsNull(rs("InvesmentID").value), "", rs("InvesmentID").value)
               If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetDataShare()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.TblIPOSharer.ID, dbo.TblIPOSharer.RecorDate, dbo.TblIPOSharer.Remark, dbo.TblIPOSharer.OrderInvse, dbo.TblIPOSharer.CountShare, "
    sql = sql & "                  dbo.TblIPOSharer.ShareValue, dbo.TblIPOSharer.ShareTotal, dbo.TblIPOSharer.ShareInvsCount, dbo.TblIPOSharer.Toatal, dbo.TblIPOSharer.TotalCountShare,"
    sql = sql & "                   dbo.TblIPOSharer.PaymentType, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblIPOSharer.BranchID, dbo.TblIPOSharer.SharID,"
    sql = sql & "                   dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
    sql = sql & "   FROM         dbo.TblIPOSharer LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblCustemers ON dbo.TblIPOSharer.SharID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblBranchesData ON dbo.TblIPOSharer.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = ""
  
    
    If val(Me.TxtFromShareValue.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblIPOSharer.ShareValue >=" & val(Me.TxtFromShareValue.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblIPOSharer.ShareValue >=" & val(Me.TxtFromShareValue.text) & ""
        End If
    End If
    If val(Me.TxtToShareValue.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblIPOSharer.ShareValue <=" & val(Me.TxtToShareValue.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblIPOSharer.ShareValue <=" & val(Me.TxtToShareValue.text) & ""
       End If
    End If
    ''''''''''''''''''''''''''''''''''''
    
       If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TblIPOSharer.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblIPOSharer.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.TblIPOSharer.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.TblIPOSharer.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
    
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblIPOSharer.RecorDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblIPOSharer.RecorDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblIPOSharer.RecorDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblIPOSharer.RecorDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcbSahr.text <> "" And (val(DcbSahr.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPOSharer.SharID =" & Me.DcbSahr.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPOSharer.SharID =" & Me.DcbSahr.BoundText & ""
       End If
     End If
        If Me.DcbBranch.text <> "" And (val(DcbBranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPOSharer.BranchID =" & Me.DcbBranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPOSharer.BranchID =" & Me.DcbBranch.BoundText & ""
       End If
     End If
       If Me.TxtCountShare.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPOSharer.CountShare =" & val(Me.TxtCountShare.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPOSharer.CountShare =" & (Me.TxtCountShare.text) & ""
       End If
     End If
         If Me.TxtShareInvsCount.text <> "" Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  TblIPOSharer.ShareInvsCount =" & val(Me.TxtShareInvsCount.text) & ""
        Else:
          BolBegine = True
          StrWhere = " Where TblIPOSharer.ShareInvsCount =" & val(Me.TxtShareInvsCount.text) & ""
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TblIPOSharer.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecorDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecorDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("ShareInvsCount")) = IIf(IsNull(rs("ShareInvsCount").value), 0, rs("ShareInvsCount").value)
                .TextMatrix(i, .ColIndex("ShareValue")) = IIf(IsNull(rs("ShareValue").value), 0, rs("ShareValue").value)
                .TextMatrix(i, .ColIndex("CountShare")) = IIf(IsNull(rs("CountShare").value), 0, rs("CountShare").value)
                
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    sql = "SELECT     dbo.Tblinvestment.ID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.Tblinvestment.InvsValue, dbo.Tblinvestment.DevlpValue, "
    sql = sql & "                  dbo.Tblinvestment.TotalInDe, dbo.Tblinvestment.AllInvsValue, dbo.Tblinvestment.warrantValue, dbo.Tblinvestment.Remark, dbo.Tblinvestment.RecorDate,"
    sql = sql & "                  dbo.Tblinvestment.StatusIPO, dbo.Tblinvestment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Tblinvestment.UserID,"
    sql = sql & "                  dbo.Tblinvestment.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.Tblinvestment.TypwInvse,"
    sql = sql & "                  dbo.TblShareType.Name AS InvTypeName, dbo.TblShareType.NameE AS InvTypeNameE, dbo.Tblinvestment.GroupInvs, dbo.TblSharesGroup.Name AS InvGrName,"
    sql = sql & "                  dbo.TblSharesGroup.NameE AS InvGrNameE"
    sql = sql & "    FROM         dbo.Tblinvestment LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblSharesGroup ON dbo.Tblinvestment.GroupInvs = dbo.TblSharesGroup.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblShareType ON dbo.Tblinvestment.TypwInvse = dbo.TblShareType.ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblEmployee ON dbo.Tblinvestment.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblBranchesData ON dbo.Tblinvestment.BranchID = dbo.TblBranchesData.branch_id"
       BolBegine = False
       StrWhere = ""
 If inde = 1 Then
      If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.Tblinvestment.StatusIPO =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.Tblinvestment.StatusIPO =1"
        End If
 End If
  If Me.TxtName.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.Tblinvestment.Name like '%" & Me.TxtName.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.Tblinvestment.Name like '%" & Me.TxtName.text & "%'"
        End If
    End If
      If Me.TxtNameE.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.Tblinvestment.NameE like '%" & Me.TxtNameE.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.Tblinvestment.NameE like '%" & Me.TxtNameE.text & "%'"
        End If
    End If
    
    
    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.Tblinvestment.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.Tblinvestment.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
          StrWhere = StrWhere & " AND dbo.Tblinvestment.ID <=" & val(Me.TxtIDTO.text) & ""
     Else
          BolBegine = True
         StrWhere = " Where dbo.Tblinvestment.ID <=" & val(Me.TxtIDTO.text) & ""
       End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.Tblinvestment.RecorDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.Tblinvestment.RecorDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.Tblinvestment.RecorDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.Tblinvestment.RecorDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcboEmpName.text <> "" And (val(DcboEmpName.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  Tblinvestment.EmpID =" & Me.DcboEmpName.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where Tblinvestment.EmpID =" & Me.DcboEmpName.BoundText & ""
       End If
     End If
        If Me.Dcbranch.text <> "" And (val(Dcbranch.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  Tblinvestment.BranchID =" & Me.Dcbranch.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where Tblinvestment.BranchID =" & Me.Dcbranch.BoundText & ""
       End If
     End If
       If Me.DcbGroupInvs.text <> "" And (val(DcbGroupInvs.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  Tblinvestment.GroupInvs =" & Me.DcbGroupInvs.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where Tblinvestment.GroupInvs =" & Me.DcbGroupInvs.BoundText & ""
       End If
     End If
         If Me.DcbTypwInvse.text <> "" And (val(DcbTypwInvse.BoundText) <> 0) Then
        If BolBegine = True Then
           StrWhere = StrWhere & " AND  Tblinvestment.TypwInvse =" & Me.DcbTypwInvse.BoundText & ""
        Else:
          BolBegine = True
          StrWhere = " Where Tblinvestment.TypwInvse =" & Me.DcbTypwInvse.BoundText & ""
       End If
     End If
     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


     ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.Tblinvestment.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lblL(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecorDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecorDate").value, "yyyy/M/d")
                End If

                
                .TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
               ' .TextMatrix(i, .ColIndex("TypeExpens")) = IIf(IsNull(rs("DetTypeExpens").value), 0, rs("DetTypeExpens").value)
               ' .TextMatrix(i, .ColIndex("Distribution")) = IIf(IsNull(rs("Distribution").value), "", rs("Distribution").value)
               '  .TextMatrix(i, .ColIndex("DetValu")) = IIf(IsNull(rs("DetValu").value), 0, rs("DetValu").value)
                
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("InvGrName")) = IIf(IsNull(rs("InvGrName").value), "", rs("InvGrName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("InvTypeName")) = IIf(IsNull(rs("InvTypeName").value), "", rs("InvTypeName").value)
                Else
                .TextMatrix(i, .ColIndex("InvGrName")) = IIf(IsNull(rs("InvGrNameE").value), "", rs("InvGrNameE").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("InvTypeName")) = IIf(IsNull(rs("InvTypeNameE").value), "", rs("InvTypeNameE").value)
               End If
    
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub

Private Sub ChangeLang()
'####################### khaled ###############################
lbl(57).Caption = "Reservation Holder"
lbl(58).Caption = "Project"
lbl(54).Caption = "Phone"
Label1(29).Caption = "Branch"
With VSFlexGrid12
    .TextMatrix(0, .ColIndex("Serial")) = "Serial"
    .TextMatrix(0, .ColIndex("id")) = "ID"
    .TextMatrix(0, .ColIndex("RecorDate")) = "Recoed Date"
    .TextMatrix(0, .ColIndex("Name")) = "Reservation Holder"
    .TextMatrix(0, .ColIndex("Project")) = "Project"
    .TextMatrix(0, .ColIndex("phone")) = "Phone"
    .TextMatrix(0, .ColIndex("Branch")) = "Branch"
End With

'##############################################################


    Cmd(1).Caption = "Clear"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
    If inde = 3 Then
    Me.Caption = " IPO Shareholders Search"
    ElseIf inde = 4 Then
    Me.Caption = " IPO Search"
    ElseIf inde = 6 Or inde = 16 Or inde = 17 Or inde = 18 Or inde = 19 Or inde = 20 Or inde = 28 Then
    Me.Caption = "Serach Active Investment"
     ElseIf inde = 7 Or inde = 8 Or inde = 71 Then
    Me.Caption = " Serach Buy Land and Real Estate "
     ElseIf inde = 10 Then
    Me.Caption = " Search Model Workflow Daily Report   "
     ElseIf inde = 11 Then
    Me.Caption = " Search Land Expenses  "
      ElseIf inde = 12 Then
    Me.Caption = " Search Divide  Land   "
      ElseIf inde = 13 Then
    Me.Caption = " Search Waiver/Sale of Shares   "
      ElseIf inde = 14 Or inde = 72 Then
    Me.Caption = " Search Sales Bill   "
         ElseIf inde = 15 Then
    Me.Caption = " Search Profit Distribution to Shareholders   "
            ElseIf inde = 27 Then
    Me.Caption = " Search Investment liquidation   "
             ElseIf inde = 30 Then
    Me.Caption = " Search Booking   "
    Else
    Me.Caption = "Investment Search"
    End If
'''//////
Label1(23).Caption = "Profit"
Label1(21).Caption = "Inv.Value"
Label1(20).Caption = "To"
Label1(22).Caption = "To"
Label1(26).Caption = "To"
Label1(24).Caption = "Investment"
Label1(28).Caption = "Total Sales"
Label1(18).Caption = "Meter Value"
Label1(19).Caption = "To"
Label1(27).Caption = "Employee"
Label1(15).Caption = "Part No"
Label1(14).Caption = "Investment"
Label1(11).Caption = "Seller"
Label1(17).Caption = "Type"
Label1(13).Caption = "Customer"
Label1(12).Caption = "Land"
lbl(50).Caption = "Type"
Label1(25).Caption = "Shareholder"
Label1(1).Caption = "Investment"
Label1(10).Caption = "Customer"
lbl(1).Caption = "Seller"
Label1(5).Caption = "Divide Type"
Label1(8).Caption = "Block No."
Label1(9).Caption = "Part No."
lbl(0).Caption = "Sahre Value "
lbl(7).Caption = "Sahre No."
lbl(17).Caption = "To"
lbl(18).Caption = "To"
Label1(7).Caption = "Investment"
Label1(6).Caption = "Land"
Label1(3).Caption = "Developer"
Label1(4).Caption = "Type Development"
lbl(39).Caption = "Des Arabic"
lbl(38).Caption = "Des English"
lbl(40).Caption = "Area"
lbl(44).Caption = "Metere Price"
lbl(43).Caption = "Payment"
lbl(41).Caption = "Plan No"
lbl(42).Caption = "Property Deed "
Label1(0).Caption = "Owner"
    ' labell name
    lbl(30).Caption = "Branch"
    lbl(29).Caption = "Manger"
    Label1(1).Caption = "Owner"
    lbl(28).Caption = "Investment No "
    lbl(26).Caption = "Investment Value "
    lbl(27).Caption = "Count Share"
    lbl(25).Caption = "Share Value"
    
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Trans ID"
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(13).Caption = "Trans Date"
    Me.lbl(4).Caption = "From"
    Me.lbl(3).Caption = "To"
    lbl(17).Caption = "Name Ar"
    lbl(18).Caption = "Name Eng"
    lbl(0).Caption = "Group Investment"
    lbl(7).Caption = "Branch"
    lbl(1).Caption = "Type Investment"
    lbl(20).Caption = "Employee"
    lbl(2).Caption = "Total"
    lbl(8).Caption = "Branch"
    lbl(15).Caption = "Shareholder"
   lbl(9).Caption = "Available Share"
   lbl(10).Caption = "Share Value From "
   lbl(11).Caption = "To"
   lbl(12).Caption = "Count Share"
   lbl(24).Caption = "Branch"
   lbl(16).Caption = "Type Investment "
   lbl(19).Caption = "Investment No "
   lbl(22).Caption = "Investment Value "
   lbl(21).Caption = "Count Share"
   lbl(23).Caption = "Count Value"
   lbl(31).Caption = "Date"
   lbl(33).Caption = "From"
   lbl(34).Caption = "To"
   lbl(32).Caption = "Branch"
   lbl(36).Caption = "Employee"
   lbl(37).Caption = "Job"
   lbl(35).Caption = "Management"
   Label1(33).Caption = "Employee"
   Label1(34).Caption = "Investment"
   lbl(51).Caption = "Liq.Date"
   lbl(52).Caption = "From"
   lbl(53).Caption = "To"
   ''''//////////
        With Me.VSFlexGrid11
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Record Date"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
        .TextMatrix(0, .ColIndex("LiqDate")) = "Liq.Date"
        .TextMatrix(0, .ColIndex("Name")) = "Investment"
    End With
    
     With Me.VSFlexGrid10
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
        .TextMatrix(0, .ColIndex("DetCusName")) = "Shareholder"
        .TextMatrix(0, .ColIndex("Name")) = "Investment"
        .TextMatrix(0, .ColIndex("InvestValue")) = "Invest.Value"
        .TextMatrix(0, .ColIndex("SalValue")) = "Sales Value"
        .TextMatrix(0, .ColIndex("PorfetValue")) = "Porfet"

    End With
    
         With Me.VSFlexGrid9
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("SellerType")) = "Type"
        .TextMatrix(0, .ColIndex("SelCusName")) = "Seller"
        .TextMatrix(0, .ColIndex("TypName")) = "Type"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer"
        .TextMatrix(0, .ColIndex("InvestName")) = "Investment"
        .TextMatrix(0, .ColIndex("Name")) = "Land"
        .TextMatrix(0, .ColIndex("PartNo")) = "Part No."
        .TextMatrix(0, .ColIndex("Valu")) = "Value"
    End With
    
      With Me.VSFlexGrid8
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("DetInvName")) = "Investment"
        .TextMatrix(0, .ColIndex("SalerCusName")) = "Seller"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer"
        .TextMatrix(0, .ColIndex("DetSharNo")) = "Shar No."
        .TextMatrix(0, .ColIndex("DetSharValue")) = "Shar Value"
    End With
    
    With Me.VSFlexGrid7
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("Name")) = "Investment"
        .TextMatrix(0, .ColIndex("LandName")) = "Land"
        .TextMatrix(0, .ColIndex("DivName")) = "Divide Type"
        .TextMatrix(0, .ColIndex("BlokNo")) = "Block No. "
        .TextMatrix(0, .ColIndex("PartNo")) = "Part No. "
    End With
    
    With Me.VSFlexGrid5
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("Name")) = "Investment"
        .TextMatrix(0, .ColIndex("LandName")) = "Land"
        .TextMatrix(0, .ColIndex("InvName")) = "Type Development"
        .TextMatrix(0, .ColIndex("CusName")) = "Developer "
    End With
    
        With Me.VSFlexGrid4
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("Name")) = "Description Arabic"
        .TextMatrix(0, .ColIndex("NameE")) = "Description English"
        .TextMatrix(0, .ColIndex("CusName")) = "Owner"
        .TextMatrix(0, .ColIndex("No_planned")) = "Plan No "
        .TextMatrix(0, .ColIndex("Area")) = "Area"
        .TextMatrix(0, .ColIndex("MeterPrice")) = "Meter Price "
        .TextMatrix(0, .ColIndex("PayType")) = "Payment "
        .TextMatrix(0, .ColIndex("TitledeedNo")) = "Property deed"
    End With
    
     With Me.VSFlexGrid3
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("DepartmentName")) = "Management"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee "
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("JobTypeName")) = "Job "
        .TextMatrix(0, .ColIndex("TransDate")) = "Date"
        .TextMatrix(0, .ColIndex("NameDay")) = "Day Name "
    End With
    
          With Me.VSFlexGrid6
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("InviseNo")) = "Investment No"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Manager Name "
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("InviseValue")) = "Investment Value"
        .TextMatrix(0, .ColIndex("SharesCount")) = "Count Share"
        .TextMatrix(0, .ColIndex("SharesValue")) = "Share Value "
        .TextMatrix(0, .ColIndex("CusName")) = "Owner Name  "
        
    End With
    ''''''''''''''''''''''' next
         With Me.VSFlexGrid2
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("InvesNo")) = "Investment No"
        .TextMatrix(0, .ColIndex("Name")) = "Type Investment "
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("InvesTotal")) = "Investment Value"
        .TextMatrix(0, .ColIndex("CountShare")) = "Count Share"
        .TextMatrix(0, .ColIndex("ShareValue")) = "Share Value "
        
    End With
    
     With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Shareholder"
        .TextMatrix(0, .ColIndex("CountShare")) = "Count Share"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("ShareValue")) = "Share Value"
        .TextMatrix(0, .ColIndex("ShareInvsCount")) = "Count Share"
    End With
    
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
        .TextMatrix(0, .ColIndex("RecorDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("Name")) = "Name Arabic"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
        .TextMatrix(0, .ColIndex("NameE")) = "Name English"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("InvGrName")) = "Group Investment"
        .TextMatrix(0, .ColIndex("InvTypeName")) = "Type Investment"
    End With
  End Sub
'''''''''''''''''''''''''''' end




Private Sub DcbEmployee_Change()
If val(DcbEmployee.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmployee.BoundText, EmpCode
    TxtDcbEmploSearch.text = EmpCode
End Sub

Private Sub DcbEmployee_Click(Area As Integer)
DcbEmployee_Change
End Sub

Private Sub NameTxt_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text1.text, 1)
End Sub

Private Sub Text29_Change()
   Dim Dcombos As New ClsDataCombos
       Dcombos.GetQuicSearch DCCar, Text29, "TblCarsData", , "BoardNO", "BoardNO"
    
End Sub

Private Sub Text29_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
               Unload FrmCasrShearches
                  ' Load FrmCasrShearches
                   FrmCasrShearches.SendForm = "FrmSearchinvestment"
                    FrmCasrShearches.show vbModal
                   End If
End Sub

Private Sub TxtDcbEmploSearch_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtDcbEmploSearch.text, EmpID
        DcbEmployee.BoundText = EmpID
    End If
End Sub
Private Sub Text11_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
Dim ID  As Integer
GetTblBuyLandRealEstate ID, Me.Text12.text, 1
DcbLand.BoundText = ID
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text13.text, EmpID
        DcbCus.BoundText = EmpID
    End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text14.text, EmpID
        DcbSales.BoundText = EmpID
    End If
End Sub



Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.text, EmpID
        DcbCusto.BoundText = EmpID
    End If
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
If val(Me.DcbTypeSales.ListIndex) = 0 Then
DcbSeler.BoundText = GeTEmpIDByEmpCode(Text16.text, True)
Else
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text16.text, EmpID
        DcbSeler.BoundText = EmpID
    End If
End If
End Sub

Private Sub Text17_KeyPress(KeyAscii As Integer)
Dim ID  As Integer
GetTblBuyLandRealEstate ID, Me.Text17.text, 1
DcbLand1.BoundText = ID
End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text18.text, EmpID
        DcbCust2.BoundText = EmpID
    End If
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text19.text, 0)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim CuID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CuID, , Text2.text
        DcbSahr.BoundText = CuID
    End If
End Sub

Private Sub Text20_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text20.text, 0)
End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Text3.text, 0)
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text22.text, 0)
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text23.text, 0)
End Sub

Private Sub Text24_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text24.text, EmpID
        DcbSharer.BoundText = EmpID
    End If
End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text25.text, 0)
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text26.text, EmpID
        DcbEmp.BoundText = EmpID
    End If
End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text27.text, 0)
End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.Text28.text, 0)
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Text3.text, 0)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Text4.text, 0)
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Text5.text, 0)
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Text6.text, 0)
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text7.text, EmpID
        Me.DcbManger.BoundText = EmpID
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text8.text, EmpID
        Me.DcbEmp10.BoundText = EmpID
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text9.text, EmpID
        DcbOwner.BoundText = EmpID
    End If
End Sub

Private Sub TxtArea_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtArea.text, 0)
End Sub

Private Sub TxtCountShare_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtCountShare.text, 1)
End Sub

Private Sub TxtFromShareValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtFromShareValue.text, 1)
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)

End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
End Sub

Private Sub TxtInvesNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtInvesNo.text, 1)
End Sub

Private Sub TxtInvesTotal_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtInvesTotal.text, 1)
End Sub

Private Sub TxtMeterValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtMeterValue.text, 0)
End Sub

Private Sub TxtName_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub dcsupplier_Click(Area As Integer)
   If val(dcsupplier.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.Text10.text = EmpCode
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text10.text, EmpID
        dcsupplier.BoundText = EmpID
    End If

End Sub

Private Sub TxtShareInvsCount_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtShareInvsCount.text, 1)
End Sub

Private Sub TxtShareValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtShareValue.text, 1)
End Sub

Private Sub TxtSharNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSharNo.text, 0)
End Sub


Private Sub TxtSharNoto_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSharNoto.text, 0)
End Sub

Private Sub TxtSharValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSharValue.text, 0)
End Sub


Private Sub TxtSharValueTo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSharValueTo.text, 0)
End Sub

Private Sub TxtToShareValue_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtToShareValue.text, 1)
End Sub

Private Sub VSFlexGrid1_Click()
If inde = 3 Then
FrmIPOSharer.FindRec val(VSFlexGrid1.TextMatrix(VSFlexGrid1.row, 1))
End If
End Sub

Private Sub VSFlexGrid10_Click()
If inde = 15 Then
FrmInvestProfitDistribution.FindRec val(VSFlexGrid10.TextMatrix(VSFlexGrid10.row, VSFlexGrid10.ColIndex("id")))
End If
End Sub

Private Sub VSFlexGrid11_Click()
If inde = 27 Then
FrmInvestliquidation.FindRec val(VSFlexGrid11.TextMatrix(VSFlexGrid11.row, 1))
End If
End Sub

Private Sub VSFlexGrid12_Click()
FrmBookingBondsInvs.FindRec val(VSFlexGrid12.TextMatrix(VSFlexGrid12.row, 1))
End Sub

Private Sub VSFlexGrid13_Click()
If inde = 31 Then
    FrmTransacRegistr.FindRec val(VSFlexGrid13.TextMatrix(VSFlexGrid13.row, 2))
End If
End Sub

Private Sub VSFlexGrid14_Click()
If inde = 32 Then
FrmTravelTransactions.Retrive val(VSFlexGrid14.TextMatrix(VSFlexGrid14.row, 2))
End If
End Sub

Private Sub VSFlexGrid2_Click()
If inde = 4 Then
FrmIPO.FindRec val(VSFlexGrid2.TextMatrix(VSFlexGrid2.row, 1))
End If
End Sub
Private Sub VSFlexGrid3_Click()
If inde = 10 Then
FrmDailyWorkflow.FindRec val(VSFlexGrid3.TextMatrix(VSFlexGrid3.row, 1))
End If
End Sub
Private Sub VSFlexGrid4_Click()
If inde = 7 Then
FrmBuylandRealEstate.FindRec val(VSFlexGrid4.TextMatrix(VSFlexGrid4.row, 1))
ElseIf inde = 8 Then
FrmActiveInvestment.DcbLand.BoundText = val(VSFlexGrid4.TextMatrix(VSFlexGrid4.row, 1))
ElseIf inde = 71 Then
FrmBuylandRealEstate.TxtBillNo.text = val(VSFlexGrid4.TextMatrix(VSFlexGrid4.row, 1))
End If
End Sub

Private Sub VSFlexGrid5_Click()
If inde = 11 Then
FrmExpensesInvestment.FindRec val(VSFlexGrid5.TextMatrix(VSFlexGrid5.row, VSFlexGrid5.ColIndex("id")))
ElseIf inde = 111 Then
FrmReturnExpensInves.FindRec val(VSFlexGrid5.TextMatrix(VSFlexGrid5.row, VSFlexGrid5.ColIndex("id")))
ElseIf inde = 110 Then
FrmReturnExpensInves.TxtOrderNo.text = val(VSFlexGrid5.TextMatrix(VSFlexGrid5.row, VSFlexGrid5.ColIndex("id")))
FrmReturnExpensInves.RetreivOrder
End If
End Sub

Private Sub VSFlexGrid6_Click()
If inde = 6 Then
FrmActiveInvestment.FindRec val(VSFlexGrid6.TextMatrix(VSFlexGrid6.row, 1))
ElseIf inde = 16 Then
FrmExpensesInvestment.DcbInvise.BoundText = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.row, VSFlexGrid6.ColIndex("InviseNo")))
ElseIf inde = 17 Then
FrmDiviInvestment.DcbInvise.BoundText = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.row, VSFlexGrid6.ColIndex("InviseNo")))
ElseIf inde = 18 Then
FrmBuyBillInvestment.DcbInvise.BoundText = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.row, VSFlexGrid6.ColIndex("InviseNo")))
ElseIf inde = 19 Then
FrmInvestProfitDistribution.DcbInvise.BoundText = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.row, VSFlexGrid6.ColIndex("InviseNo")))
ElseIf inde = 20 Then
FrmInvesSales.DcbInvise.BoundText = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.row, VSFlexGrid6.ColIndex("InviseNo")))
ElseIf inde = 28 Then
FrmInvestliquidation.DcbInvest.BoundText = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.row, VSFlexGrid6.ColIndex("InviseNo")))
End If
End Sub

Private Sub VSFlexGrid7_Click()
If inde = 12 Then
FrmDiviInvestment.FindRec val(VSFlexGrid7.TextMatrix(VSFlexGrid7.row, VSFlexGrid7.ColIndex("id")))
End If
End Sub

Private Sub VSFlexGrid8_Click()
If inde = 13 Then
FrmBuyBillInvestment.FindRec val(VSFlexGrid8.TextMatrix(VSFlexGrid8.row, VSFlexGrid8.ColIndex("id")))
End If
End Sub

Private Sub VSFlexGrid9_Click()
If inde = 14 Then
FrmSaleBillInvestment.FindRec val(VSFlexGrid9.TextMatrix(VSFlexGrid9.row, VSFlexGrid9.ColIndex("id")))
ElseIf inde = 72 Then
FrmSaleBillInvestment.TxtBillNo.text = val(VSFlexGrid9.TextMatrix(VSFlexGrid9.row, VSFlexGrid9.ColIndex("id")))
End If
End Sub
Private Sub projTxtSearchCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
DcbProject.BoundText = projTxtSearchCode.text
End If
End Sub
Private Sub DcbProject_Change()
DcbProject_Click (0)
End Sub

Private Sub DcbProject_Click(Area As Integer)
If DcbProject.BoundText <> "" Then
projTxtSearchCode.text = DcbProject.BoundText
End If
End Sub
Public Sub GetKData()

    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    sql = "SELECT TblBookingBondsInvs.ID, TblBookingBondsInvs.BrnchID, TblBookingBondsInvs.RecordDate, TblBookingBondsInvs.Name, TblBookingBondsInvs.Telephone, TblProjecInvestment.Name AS ProjectName, TblProjecInvestment.NameE, "
    sql = sql & " TblBranchesData.branch_name , TblBranchesData.branch_namee "
    sql = sql & " FROM TblBookingBondsInvs INNER JOIN "
    sql = sql & " TblProjecInvestment ON TblBookingBondsInvs.ProjectID = TblProjecInvestment.ID INNER JOIN "
    sql = sql & " TblBranchesData ON TblBookingBondsInvs.BrnchID = TblBranchesData.branch_id "
    sql = sql & " WHERE  (1=1) "
    
    '############################################################################################################################
    If val(TxtIDFrom.text) <> 0 Then
        sql = sql & " AND  dbo.TblBookingBondsInvs.ID >=" & val(Me.TxtIDFrom.text) & ""
    End If
    If Me.TxtIDTO.text <> "" Then
          sql = sql & " AND dbo.TblBookingBondsInvs.ID <=" & val(Me.TxtIDTO.text) & ""
    End If
    '############################################################################################################################
    If Not IsNull(Me.DtpDateFrom.value) Then
            sql = sql & " AND dbo.TblBookingBondsInvs.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
            sql = sql & " AND dbo.TblBookingBondsInvs.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If
    '############################################################################################################################
    If Me.TxtNameP.text <> "" Then
           sql = sql & " AND  dbo.TblBookingBondsInvs.Name Like N'%" & (Me.TxtNameP.text) & "%'"
    End If
    '############################################################################################################################
    If Me.DcbProject.text <> "" And (val(DcbProject.BoundText) <> 0) Then
           sql = sql & " AND  dbo.TblBookingBondsInvs.ProjectID =" & val(Me.DcbProject.BoundText) & ""
    End If
    '############################################################################################################################
    If Me.TxtTelephone.text <> "" Then
           sql = sql & " AND  dbo.TblBookingBondsInvs.Telephone Like N'%" & (Me.TxtTelephone.text) & "%'"
    End If
    '############################################################################################################################
    If Me.DcbKBranch.text <> "" And (val(DcbKBranch.BoundText) <> 0) Then
           sql = sql & " AND  dbo.TblBookingBondsInvs.BrnchID =" & val(Me.DcbKBranch.BoundText) & ""
    End If
    '############################################################################################################################
    
    sql = sql & " Order By dbo.TblBookingBondsInvs.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        End If
        With Me.VSFlexGrid12
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
        End With
        Exit Sub
    Else
        With Me.VSFlexGrid12
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            'If SystemOptions.UserInterface = ArabicInterface Then
            '    Me.lblL(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            'ElseIf SystemOptions.UserInterface = EnglishInterface Then
            '    Me.lblL(10).Caption = "Search Results=" & rs.RecordCount
            'End If
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("Phone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Project")) = IIf(IsNull(rs("ProjectName").value), "", rs("ProjectName").value)
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    .TextMatrix(i, .ColIndex("Project")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetArchData()

    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    sql = " SELECT TblTransacRegistr.ID, TblTransacRegistr.BrnchID, TblTransacRegistr.RecordDate, TblTransacRegistr.RecordDateH, TblTransacRegistr.RecordTime, TblTransacRegistr.UserID, TblTransacRegistr.barcode, "
    sql = sql & " TblTransacRegistr.TypTrans, TblTransacRegistr.ImportExport, TblTransacRegistr.NoImpExp, TblTransacRegistr.ImpExpDate, TblTransacRegistr.ImpExpDateH, TblTransacRegistr.Summary, TblTransacRegistr.EnterDate, "
    sql = sql & " TblTransacRegistr.EnterTime, TblTransacRegistr.RequerTime, TblTransacRegistr.ExitTime, TblTransacRegistr.ProcedureReq, TblTransacRegistr.Remarks, TblTransacRegistr.ExitDate, TblTransacRegistr.MHD, "
    sql = sql & " TblTransacRegistr.MHDID, TblTransacRegistr.Posted, TblTransacRegistr.PostedDate, TblTransacRegistr.Approved, TblTransacRegistrDet.TransRegID, TblTransacRegistrDet.FromUser, "
    sql = sql & " TblTransacRegistrDet.ID AS TblTransacRegistrDetID, TblTransacRegistrDet.ToUser, TblTransacRegistrDet.FlgTrans, TblTransacRegistrDet.RecDate, TblTransacRegistrDet.ProcedureReq AS ProcedureReqDet, "
    sql = sql & " TblTransacRegistrDet.Time , TblTransacRegistrDet.TimeUnitID, TblBranchesData.branch_name, TblBranchesData.branch_namee ,TblTransacRegistr.NoteSerial1"
    sql = sql & " FROM TblTransacRegistr INNER JOIN "
    sql = sql & " TblBranchesData ON TblTransacRegistr.BrnchID = TblBranchesData.branch_id FULL OUTER JOIN "
    sql = sql & " TblTransacRegistrDet ON TblTransacRegistr.ID = TblTransacRegistrDet.TransRegID "
    sql = sql & " Where (1 = 1) "
    
    '############################################################################################################################
    If val(TxtIDFrom.text) <> 0 Then
        sql = sql & " AND  TblTransacRegistr.NoteSerial1 > =" & val(Me.TxtIDFrom.text) & ""
    End If
    If Me.TxtIDTO.text <> "" Then
          sql = sql & " AND TblTransacRegistr.NoteSerial1 < =" & val(Me.TxtIDTO.text) & ""
    End If
    '############################################################################################################################
    If Not IsNull(Me.DtpDateFrom.value) Then
            sql = sql & " AND TblTransacRegistr.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
            sql = sql & " AND TblTransacRegistr.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If
    '############################################################################################################################
    If Me.TxtNoImpExp.text <> "" Then
           sql = sql & " AND  TblTransacRegistr.NoImpExp Like N'%" & (Me.TxtNoImpExp.text) & "%'"
    End If
    '############################################################################################################################
    If Me.Txtbarcode.text <> "" Then
           sql = sql & " AND  TblTransacRegistr.barcode Like N'%" & (Me.Txtbarcode.text) & "%'"
    End If
    '############################################################################################################################
    If Me.ArchSearchBranchDC.text <> "" And (val(ArchSearchBranchDC.BoundText) <> 0) Then
           sql = sql & " AND  TblTransacRegistr.BrnchID = " & val(Me.ArchSearchBranchDC.BoundText) & ""
    End If
    '############################################################################################################################
    If Me.DcbImportExport.text <> "" And (val(DcbImportExport.ListIndex) <> -1) Then
           sql = sql & " AND  TblTransacRegistr.ImportExport = " & val(Me.DcbImportExport.ListIndex) & ""
    End If
    '############################################################################################################################
        If Me.Summary.text <> "" Then
           sql = sql & " AND  TblTransacRegistr.Summary Like N'%" & (Me.Summary.text) & "%'"
    End If
    '############################################################################################################################
    If Not IsNull(Me.FrmEnterDate.value) Then
            sql = sql & " AND TblTransacRegistr.EnterDate >=" & SQLDate(Me.FrmEnterDate.value, True) & ""
    End If
    If Not IsNull(Me.ToEnterDate.value) Then
            sql = sql & " AND TblTransacRegistr.EnterDate <=" & SQLDate(Me.ToEnterDate.value, True) & ""
    End If
    '############################################################################################################################
        If Not IsNull(Me.FrmExitDate.value) Then
            sql = sql & " AND TblTransacRegistr.ExitDate >=" & SQLDate(Me.FrmExitDate.value, True) & ""
    End If
    If Not IsNull(Me.ToExitDate.value) Then
            sql = sql & " AND TblTransacRegistr.ExitDate <=" & SQLDate(Me.ToExitDate.value, True) & ""
    End If
    '############################################################################################################################
    
    sql = sql & " Order By dbo.TblTransacRegistr.ID "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        End If
        With Me.VSFlexGrid13
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
        End With
        Exit Sub
    Else
        With Me.VSFlexGrid13
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_name").value)
                End If
                If Not (IsNull(rs("EnterDate").value)) Then
                    .TextMatrix(i, .ColIndex("EnterDate")) = Format(rs("EnterDate").value, "yyyy/M/d")
                End If
                If Not (IsNull(rs("ExitDate").value)) Then
                    .TextMatrix(i, .ColIndex("ExitDate")) = Format(rs("ExitDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("NoImpExp")) = IIf(IsNull(rs("NoImpExp").value), "", rs("NoImpExp").value)
                .TextMatrix(i, .ColIndex("barcode")) = IIf(IsNull(rs("barcode").value), "", rs("barcode").value)
                If IsNull(rs("barcode").value) Then
                Else
                    If rs("TypTrans").value = 0 Then
                        .TextMatrix(i, .ColIndex("TypTrans")) = "’«œ—"
                    ElseIf rs("TypTrans").value = 1 Then
                     .TextMatrix(i, .ColIndex("TypTrans")) = "Ê«—œ"
                    End If
                End If
                .TextMatrix(i, .ColIndex("Summary")) = IIf(IsNull(rs("Summary").value), "", rs("Summary").value)
               rs.MoveNext
          Next i
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
Public Sub GetTripData()

    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim TpeiD As Integer
    sql = " SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial1, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, "
    sql = sql & "                   dbo.TblBranchesData.branch_namee, dbo.notes_all.ManualNo, dbo.notes_all.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    sql = sql & "                   dbo.TblCustemers.Fullcode, dbo.notes_all.ShipID, dbo.TblShipsData.Name, dbo.TblShipsData.NameE, dbo.notes_all.CityFromId,"
    sql = sql & "                   dbo.TblCountriesGovernments.GovernmentName, dbo.notes_all.CityToId, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
    sql = sql & "                   dbo.notes_all.VehicleType, dbo.TBLCarTypes.name AS CarNameType, dbo.TBLCarTypes.namee AS CarNameTypeE, dbo.notes_all.CarType, dbo.notes_all.CarId,"
    sql = sql & "                   dbo.TblCarsData.BoardNO, dbo.notes_all.CarID2, notes_all.ContainerNo, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.notes_all.SupplemID, dbo.FixedAssets.Name AS PartName,"
    sql = sql & "                   dbo.notes_all.SupplemID2, TblVendorCars_1.accessory AS PartName2, dbo.notes_all.LeaderName, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_Name,"
    sql = sql & "                   dbo.TblEmployee.Fullcode AS EmpCode, dbo.TblEmployee.Emp_Namee"
    sql = sql & "            FROM         dbo.notes_all LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblVendorCars TblVendorCars_1 ON dbo.notes_all.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
    sql = sql & "                   dbo.FixedAssets ON dbo.notes_all.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblVendorCars ON dbo.notes_all.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TBLCarTypes ON dbo.notes_all.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblCountriesGovernments ON dbo.notes_all.CityFromId = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblShipsData ON dbo.notes_all.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
    sql = sql & "      Where (dbo.notes_all.NoteType = 370)"
    
    '############################################################################################################################
    If val(TxtIDFrom.text) <> 0 Then
        sql = sql & " AND  notes_all.NoteSerial1 > =" & val(Me.TxtIDFrom.text) & ""
    End If
    If Me.TxtIDTO.text <> "" Then
          sql = sql & " AND notes_all.NoteSerial1 < =" & val(Me.TxtIDTO.text) & ""
    End If
    '############################################################################################################################
    If Not IsNull(Me.DtpDateFrom.value) Then
            sql = sql & " AND notes_all.NoteDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
            sql = sql & " AND notes_all.NoteDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
    End If
    '############################################################################################################################
    If Me.TxtManualNo.text <> "" Then
           sql = sql & " AND  notes_all.ManualNo Like N'%" & (Me.TxtManualNo.text) & "%'"
    End If
    '############################################################################################################################
    If Me.TxtLeaderName.text <> "" Then
           sql = sql & " AND  notes_all.LeaderName Like N'%" & (Me.TxtLeaderName.text) & "%'"
    End If
    
    
   If Me.txtContainerNo.text <> "" Then
           sql = sql & " AND  notes_all.ContainerNo Like N'" & (Me.txtContainerNo.text) & "'"
    End If
     
    
    '############################################################################################################################
    If DcbBranch22.text <> "" And val(DcbBranch22.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.branch_no = " & val(DcbBranch22.BoundText) & ""
    End If
    '############################################################################################################################
    If Me.DBCboClientName.text <> "" And val(DBCboClientName.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.CusID = " & val(Me.DBCboClientName.BoundText) & ""
    End If
    '############################################################################################################################
    If Me.DcbShip.text <> "" And val(DcbShip.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.ShipID = " & val(Me.DcbShip.BoundText) & ""
    End If
    If Me.DcCityFromId.text <> "" And val(DcCityFromId.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.CityFromId = " & val(Me.DcCityFromId.BoundText) & ""
    End If
    If Me.DcCityToId.text <> "" And val(DcCityToId.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.CityToId = " & val(Me.DcCityToId.BoundText) & ""
    End If
    If Me.VehicleType.text <> "" And val(VehicleType.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.VehicleType = " & val(Me.VehicleType.BoundText) & ""
    End If
    If Me.DCCar.text <> "" And val(DCCar.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.CarId = " & val(Me.DCCar.BoundText) & ""
    End If
    If Me.DcbCar2.text <> "" And val(DcbCar2.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.CarID2 = " & val(Me.DcbCar2.BoundText) & ""
    End If
    If Me.DcbSupplem.text <> "" And val(DcbSupplem.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.SupplemID = " & val(Me.DcbSupplem.BoundText) & ""
    End If
    If Me.DcbSupplem2.text <> "" And val(DcbSupplem2.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.SupplemID2 = " & val(Me.DcbSupplem2.BoundText) & ""
    End If
    If Me.DCEmp.text <> "" And val(DCEmp.BoundText) <> 0 Then
           sql = sql & " AND  notes_all.DriverId = " & val(Me.DCEmp.BoundText) & ""
    End If
    If ChCarType(0).value = True Then
           sql = sql & " AND  (notes_all.CarType = 0 or notes_all.CarType is null)"
    End If
     If ChCarType(1).value = True Then
           sql = sql & " AND  notes_all.CarType =1"
    End If
    '############################################################################################################################
    
    sql = sql & " Order By dbo.notes_all.NoteSerial1 "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Else
         MsgBox "Sorry...no Data ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        End If
        With Me.VSFlexGrid14
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
        End With
        Exit Sub
    Else
        With Me.VSFlexGrid14
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            rs.MoveFirst
                 For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
                If Not (IsNull(rs("NoteDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecorDate")) = Format(rs("NoteDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CarNameType")) = IIf(IsNull(rs("CarNameType").value), "", rs("CarNameType").value)
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    .TextMatrix(i, .ColIndex("CarNameType")) = IIf(IsNull(rs("CarNameTypeE").value), "", rs("CarNameTypeE").value)
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_name").value)
                End If
                .TextMatrix(i, .ColIndex("ContainerNo")) = IIf(IsNull(rs("ContainerNo").value), "", rs("ContainerNo").value)
                .TextMatrix(i, .ColIndex("ManualNo")) = IIf(IsNull(rs("ManualNo").value), "", rs("ManualNo").value)
                .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
                .TextMatrix(i, .ColIndex("ToGovernmentName")) = IIf(IsNull(rs("ToGovernmentName").value), "", rs("ToGovernmentName").value)
                TpeiD = IIf(IsNull(rs("CarType").value), -1, rs("CarType").value)
                If TpeiD = 0 Or TpeiD = -1 Then
                .TextMatrix(i, .ColIndex("Car")) = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
                .TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs("PartName").value), "", rs("PartName").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Car")) = " „„·Êﬂ… ··‘—ﬂ…"
                Else
                .TextMatrix(i, .ColIndex("Car")) = "Owned"
                End If
                ElseIf TpeiD = 1 Then
                .TextMatrix(i, .ColIndex("Car")) = IIf(IsNull(rs("BoardNo2").value), "", rs("BoardNo2").value)
                .TextMatrix(i, .ColIndex("Part")) = IIf(IsNull(rs("PartName2").value), "", rs("PartName2").value)
                .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Car")) = "€Ì— „„·Êﬂ… ··‘—ﬂ…"
                Else
                .TextMatrix(i, .ColIndex("Car")) = "Not Owned"
                End If
                Else
                .TextMatrix(i, .ColIndex("Car")) = ""
                .TextMatrix(i, .ColIndex("Part")) = ""
                .TextMatrix(i, .ColIndex("DriverName")) = ""
                End If
               rs.MoveNext
          Next i
         
            .AutoSize 0, .Cols - 1, False
            
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With
    End If
End Sub
