VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmExpenses40E 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Œ·’ «Ê «” »⁄«œ«  «·«’Ê·"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
   HelpContextID   =   280
   Icon            =   "FrmExpenses40E.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   9960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.PictureBox Picture2 
      Height          =   45
      Left            =   6090
      ScaleHeight     =   45
      ScaleWidth      =   3465
      TabIndex        =   208
      Top             =   8370
      Width           =   3465
   End
   Begin VB.TextBox XPTxtValView 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   -2520
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   114
      Top             =   8040
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   8175
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   720
      Width           =   9975
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„— "
         Height          =   2055
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   172
         Top             =   840
         Width           =   4815
         Begin VB.TextBox TxtMonthVAT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox TxtNetMonth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox TxtFATValue 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   1680
            Width           =   765
         End
         Begin VB.TextBox TxtBillVAT 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   1320
            Width           =   765
         End
         Begin VB.TextBox TxtAgeMonth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   600
            Width           =   765
         End
         Begin VB.OptionButton RdMove 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " €Ì— „‰ﬁÊ· «·„œ…"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton RdMove 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ﬁÊ· «·„œ…"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtYearMove 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   240
            Width           =   765
         End
         Begin VB.TextBox TxtYearNotMove 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   181
            TabStop         =   0   'False
            Top             =   270
            Width           =   765
         End
         Begin VB.TextBox TxtMonthMove 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            TabIndex        =   180
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox TxtMonthNotMove 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.TextBox TxtPeriodMonth 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   960
            Width           =   765
         End
         Begin VB.TextBox TxtPeriodYear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox TxtAgeYear 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   600
            Width           =   855
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Bindings        =   "FrmExpenses40E.frx":038A
            Height          =   315
            Left            =   0
            TabIndex        =   194
            Top             =   0
            Visible         =   0   'False
            Width           =   3450
            _ExtentX        =   6085
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSComCtl2.DTPicker DpPurchaseDate 
            Height          =   345
            Left            =   360
            TabIndex        =   195
            Top             =   1680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   217513985
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·‘—«¡"
            Height          =   255
            Index           =   42
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·‘Â—"
            Height          =   255
            Left            =   2010
            RightToLeft     =   -1  'True
            TabIndex        =   198
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„»·€ «·›« "
            Height          =   375
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·›«  ›Ì «·‘—«¡"
            Height          =   375
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   255
            Index           =   41
            Left            =   360
            TabIndex        =   189
            Top             =   960
            Width           =   420
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‘Â—"
            Height          =   255
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„œ… «·„ »ﬁÌ…"
            Height          =   375
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   187
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   255
            Index           =   44
            Left            =   360
            TabIndex        =   183
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‘Â—"
            Height          =   255
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‘Â—"
            Height          =   255
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„— «·÷—Ì»Ì"
            Height          =   375
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   600
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· Œ·’ «·Ã“∆Ì „‰ «’·  ( ﬁ”Ì„ «·«’·)"
         Height          =   1635
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   156
         Top             =   4080
         Width           =   4935
         Begin VB.TextBox txtCountF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2430
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   900
            Width           =   1095
         End
         Begin VB.TextBox TxtNewNuminstallmRemin 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   164
            Top             =   1230
            Width           =   735
         End
         Begin VB.TextBox TxtNewCurrentValue 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   1230
            Width           =   1095
         End
         Begin VB.TextBox TxtNewAssest 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   157
            Top             =   240
            Width           =   3435
         End
         Begin MSDataListLib.DataCombo DcbGroup 
            Height          =   315
            Left            =   120
            TabIndex        =   159
            Top             =   600
            Width           =   3435
            _ExtentX        =   6059
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄œœ"
            Height          =   285
            Index           =   48
            Left            =   3450
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   900
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„ »ﬁÌ…"
            Height          =   405
            Index           =   40
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   1230
            Width           =   1395
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·œ› —Ì…"
            Height          =   285
            Index           =   39
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   1230
            Width           =   1395
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ã„Ê⁄Â"
            Height          =   375
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·«’· «·ÃœÌœ"
            Height          =   375
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·«” »⁄«œ"
         Height          =   1215
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   142
         Top             =   2880
         Width           =   4815
         Begin VB.TextBox TxtExcludedValueInst 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TxtExcludedValueFixed 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   810
            Width           =   735
         End
         Begin VB.TextBox TxtExcludedValuePrt 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TxtExcludedValue 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   480
            Width           =   735
         End
         Begin VB.OptionButton RdExcluded 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” »⁄«œ ‰”»Â"
            Height          =   255
            Index           =   1
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   144
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton RdExcluded 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«” »⁄«œ ﬁÌ„Â"
            Height          =   255
            Index           =   0
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„Â  «·ﬁ”ÿ «·ÃœÌœ…"
            Height          =   375
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„Â  «·«’· «·„ »ﬁÌ"
            Height          =   375
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„Â  «·Ã“¡ «·„” »⁄œ"
            Height          =   375
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„Â / ‰”»Â «·«” »⁄«œ"
            Height          =   375
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·»Ì⁄"
         Height          =   1785
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   132
         Top             =   5700
         Width           =   4815
         Begin VB.TextBox txtNet 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   660
            Width           =   1995
         End
         Begin VB.TextBox txtVat2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   180
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox TxtLoseProfitValue 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   360
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   170
            Top             =   990
            Width           =   2715
         End
         Begin VB.TextBox TxtFASalesPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2430
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   169
            Top             =   240
            Width           =   1245
         End
         Begin MSDataListLib.DataCombo DCAccounts 
            Height          =   315
            Left            =   360
            TabIndex        =   171
            Top             =   1350
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ï »⁄œ «·ﬁÌ„… «·„÷«›…"
            Height          =   270
            Index           =   47
            Left            =   2610
            TabIndex        =   205
            Top             =   660
            Width           =   1800
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„Â «·÷—Ì»…"
            Height          =   285
            Index           =   46
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„Â «·»Ì⁄"
            Height          =   285
            Index           =   31
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   240
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·—»Õ Ê «·Œ”«—…"
            Height          =   405
            Index           =   26
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   1350
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„Â «·—»Õ «Ê «·Œ”«—…"
            Height          =   405
            Index           =   32
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   990
            Width           =   1515
         End
      End
      Begin VB.ComboBox CboType2 
         Height          =   315
         ItemData        =   "FrmExpenses40E.frx":039F
         Left            =   360
         List            =   "FrmExpenses40E.frx":03A1
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   130
         Top             =   480
         Width           =   3375
      End
      Begin VB.ComboBox CboType 
         Height          =   315
         ItemData        =   "FrmExpenses40E.frx":03A3
         Left            =   5040
         List            =   "FrmExpenses40E.frx":03A5
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   124
         Top             =   480
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·«’· «·Õ«·Ì"
         Height          =   3135
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   117
         Top             =   1200
         Width           =   4815
         Begin VB.TextBox txtAdditions 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   200
            Top             =   960
            Width           =   2955
         End
         Begin VB.TextBox TxtNuminstallmCurr 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   153
            Top             =   2760
            Width           =   2955
         End
         Begin VB.TextBox TxtNuminstallmTotal 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   141
            Top             =   1680
            Width           =   2955
         End
         Begin VB.TextBox TxtNuminstallmRemin 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   139
            Top             =   2400
            Width           =   2955
         End
         Begin VB.TextBox TxtNuminstallmExcu 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   137
            Top             =   2040
            Width           =   2955
         End
         Begin VB.TextBox TxtCurrentValue 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   123
            Top             =   1350
            Width           =   2955
         End
         Begin VB.TextBox TxtAccDepre 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   122
            Top             =   600
            Width           =   2955
         End
         Begin VB.TextBox TxtPurchasePrice 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   121
            Top             =   240
            Width           =   2955
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«÷«›« "
            Height          =   285
            Index           =   45
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   201
            Top             =   960
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„Â «·ﬁ”ÿ «·Õ«·Ì…"
            Height          =   405
            Index           =   38
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   2760
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·«Ã„«·Ì…"
            Height          =   285
            Index           =   37
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   1680
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„ »ﬁÌ…"
            Height          =   285
            Index           =   36
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   2400
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·«ﬁ”«ÿ «·„‰›–…"
            Height          =   285
            Index           =   35
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   2040
            Width           =   1635
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·œ› —Ì…"
            Height          =   285
            Index           =   30
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1320
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã„⁄ «·«Â·«ﬂ"
            Height          =   285
            Index           =   29
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   600
            Width           =   1515
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„Â «·‘—«¡"
            Height          =   285
            Index           =   28
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   240
            Width           =   1515
         End
      End
      Begin VB.TextBox TxtNoteSerial 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -720
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   4080
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Text            =   "Text1"
         Top             =   1830
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Frame FraNote 
         BackColor       =   &H00E2E9E9&
         Height          =   3285
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   4080
         Width           =   4755
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   165
            Top             =   240
            Width           =   3315
         End
         Begin VB.TextBox TXTBankName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   1320
            Visible         =   0   'False
            Width           =   3285
         End
         Begin VB.TextBox TxtChequeNumber 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   1680
            Width           =   3285
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   600
            Width           =   705
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1320
            Width           =   705
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   960
            Width           =   705
         End
         Begin MSComCtl2.DTPicker DtpChequeDueDate 
            Height          =   315
            Left            =   150
            TabIndex        =   85
            Top             =   2460
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Format          =   219938817
            CurrentDate     =   39614
         End
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   150
            TabIndex        =   86
            Top             =   1320
            Width           =   2685
            _ExtentX        =   4736
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   120
            TabIndex        =   87
            Top             =   960
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCVendor 
            Height          =   315
            Left            =   120
            TabIndex        =   88
            Top             =   600
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCAccounts1 
            Height          =   315
            Left            =   120
            TabIndex        =   125
            Top             =   2760
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcChequeBox 
            Height          =   315
            Left            =   150
            TabIndex        =   127
            Top             =   2040
            Width           =   3285
            _ExtentX        =   5794
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·»Ì⁄"
            Height          =   255
            Index           =   15
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«›Ÿ… «·‘Ìﬂ« "
            Height          =   285
            Index           =   43
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ”«»"
            Height          =   285
            Index           =   33
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   2880
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·Œ“Ì‰…"
            Height          =   285
            Index           =   16
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   1020
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·»‰ﬂ"
            Height          =   285
            Index           =   17
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1350
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·‘Ìﬂ"
            Height          =   285
            Index           =   18
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
            Height          =   285
            Index           =   19
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   2460
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   285
            Index           =   22
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   660
            Width           =   1215
         End
      End
      Begin VB.TextBox txtto 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   15480
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   79
         Top             =   4560
         Width           =   2715
      End
      Begin VB.TextBox TxtSerial1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6720
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   150
         Width           =   1455
      End
      Begin VB.TextBox txt_general_des 
         Alignment       =   1  'Right Justify
         Height          =   525
         Left            =   240
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   77
         Top             =   7530
         Width           =   2715
      End
      Begin VB.TextBox txt_ORDER_NO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   15360
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   1590
         Width           =   2655
      End
      Begin VB.ComboBox CboPaymentType1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "FrmExpenses40E.frx":03A7
         Left            =   10920
         List            =   "FrmExpenses40E.frx":03A9
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   75
         Top             =   510
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   -240
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   1590
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox Text7 
         DataField       =   "id"
         DataSource      =   "Adodc4"
         Height          =   285
         Left            =   960
         TabIndex        =   72
         Text            =   "Text2"
         Top             =   1830
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TXT_A_NoteID 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Text            =   "Text8"
         Top             =   3270
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   4800
         TabIndex        =   96
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Format          =   220463105
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   7
         Left            =   -90
         TabIndex        =   97
         Top             =   3390
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "«·⁄—÷ «·ÃœÊ·Ï"
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   4210752
      End
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   16080
         TabIndex        =   98
         Top             =   1140
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcCostCenter 
         Bindings        =   "FrmExpenses40E.frx":03AB
         Height          =   315
         Left            =   16440
         TabIndex        =   99
         Top             =   780
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmExpenses40E.frx":03C0
         Height          =   315
         Left            =   360
         TabIndex        =   112
         Top             =   120
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
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
      Begin MSDataListLib.DataCombo DcFixedAssets 
         Height          =   315
         Left            =   5040
         TabIndex        =   115
         Top             =   840
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·⁄„·Ì…"
         Height          =   285
         Index           =   34
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   131
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·«’·"
         Height          =   285
         Index           =   27
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   116
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   113
         Top             =   120
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”‰œ"
         Height          =   285
         Index           =   4
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   110
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·„’—Ê›« "
         Height          =   285
         Index           =   3
         Left            =   15480
         RightToLeft     =   -1  'True
         TabIndex        =   109
         Top             =   1110
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Top             =   135
         Width           =   675
      End
      Begin VB.Image ImgNote 
         Height          =   240
         Left            =   -240
         Picture         =   "FrmExpenses40E.frx":03D5
         Top             =   750
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘—Ê⁄"
         Height          =   255
         Index           =   14
         Left            =   14400
         RightToLeft     =   -1  'True
         TabIndex        =   107
         Top             =   1140
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ ›« Ê—… «·‘—«¡"
         Height          =   285
         Index           =   0
         Left            =   15240
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„—ﬂ“ «· ﬂ·›… «·⁄«„"
         Height          =   255
         Left            =   15120
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   810
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "”»» «· Œ·’"
         Height          =   285
         Index           =   20
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   7650
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»Ì…"
         Height          =   285
         Index           =   21
         Left            =   15240
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   1590
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·”‰œ"
         Height          =   285
         Index           =   23
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   510
         Width           =   1275
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   735
         Left            =   16080
         Top             =   5040
         Width           =   1695
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ… Â«„…:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   24
         Left            =   16200
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   4920
         Width           =   1275
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
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
         Height          =   540
         Index           =   25
         Left            =   15480
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   4920
         Width           =   1695
      End
   End
   Begin VB.OptionButton OptSort 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   1
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   240
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox ChkLastAccount 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   240
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "Option1"
      Height          =   195
      Index           =   0
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   2340
      Left            =   16560
      TabIndex        =   47
      Top             =   4440
      Visible         =   0   'False
      Width           =   10755
      _cx             =   18971
      _cy             =   4128
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   18
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExpenses40E.frx":095F
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
      ExplorerBar     =   0
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
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   3915
         Left            =   2550
         RightToLeft     =   -1  'True
         ScaleHeight     =   3915
         ScaleWidth      =   9405
         TabIndex        =   52
         Top             =   810
         Visible         =   0   'False
         Width           =   9405
         Begin VB.CommandButton Command3 
            Caption         =   "Call des"
            Height          =   255
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   3600
            Width           =   1095
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Add des"
            Height          =   255
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   3600
            Width           =   1350
         End
         Begin VB.TextBox txtcodesub 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3600
            Width           =   855
         End
         Begin VB.TextBox TxtDese 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1485
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   53
            Top             =   2040
            Width           =   8955
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   3900
            Left            =   120
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   0
            Width           =   10905
            _cx             =   19235
            _cy             =   6879
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   20.25
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
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   7
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
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               Height          =   1605
               Left            =   0
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   3  'Both
               TabIndex        =   58
               Top             =   480
               Visible         =   0   'False
               Width           =   8955
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000C&
               Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
               ForeColor       =   &H0000C8FF&
               Height          =   315
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   0
               Width           =   2445
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   255
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Code"
            Height          =   495
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   3480
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Õœœ —ﬁ„ «·ﬁÌœ «·„—«œ ‰”Œ…"
         Height          =   1215
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   3720
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            Caption         =   "‰”Œ"
            Height          =   255
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   255
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   1335
         End
      End
      Begin VDSCOMBOLibCtl.SmartCombo SmartCombo1 
         Height          =   315
         Left            =   240
         TabIndex        =   63
         ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
         Top             =   480
         Visible         =   0   'False
         Width           =   2475
         _cx             =   1973752078
         _cy             =   1973748268
         Alignment       =   0
         Appearance      =   3
         AutoSearch      =   0   'False
         BackColor       =   -2147483624
         BackgroundColor =   -2147483633
         BorderColor     =   0
         BorderVisible   =   -1  'True
         Caption         =   "SmartCombo1"
         CaptionAlignment=   4
         CaptionBackColor=   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   -2147483630
         CaptionHeight   =   15
         CaptionOnTop    =   0   'False
         CaptionMultiLine=   0
         Checkbox3D      =   0   'False
         CheckboxAlignment=   5
         CheckboxBackColor=   16777215
         CheckboxSize    =   13
         CheckboxValue   =   0
         BrowsePictureAlignment=   5
         BrowsePictureStretchH=   0
         BrowsePictureStretchV=   0
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
         ForeColor       =   0
         Gap             =   0
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   0
         MultiLine       =   0
         OnFocus         =   3
         PasswordChar    =   ""
         Picture         =   "FrmExpenses40E.frx":0C3B
         PictureAlignment=   5
         PictureBackColor=   -2147483624
         PictureStretchH =   0
         PictureStretchV =   0
         Redraw          =   -1  'True
         ScrollBar       =   0
         Style           =   0
         Text            =   ""
         UnderLine       =   0   'False
         Enabled0        =   -1  'True
         Position0       =   0
         Tip0            =   "Caption"
         Visible0        =   0   'False
         Width0          =   90
         Enabled1        =   -1  'True
         Position1       =   1
         Tip1            =   ""
         Visible1        =   -1  'True
         Width1          =   32
         Enabled2        =   -1  'True
         Position2       =   2
         Tip2            =   "Check Box (Space, Ctrl + Space)"
         Visible2        =   0   'False
         Width2          =   16
         Enabled3        =   -1  'True
         Position3       =   3
         Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
         Visible3        =   -1  'True
         Width3          =   113
         Enabled4        =   -1  'True
         Position4       =   4
         Tip4            =   "Left Spinner (Alt + Left)"
         Visible4        =   0   'False
         Width4          =   16
         Enabled5        =   -1  'True
         Position5       =   5
         Tip5            =   "Right Spinner (Alt + Right)"
         Visible5        =   0   'False
         Width5          =   16
         Enabled6        =   -1  'True
         Position6       =   6
         Tip6            =   "Up Spinner (Ctrl + Up)"
         Visible6        =   0   'False
         Width6          =   16
         Enabled7        =   -1  'True
         Position7       =   7
         Tip7            =   "Down Spinner (Ctrl + Down)"
         Visible7        =   0   'False
         Width7          =   16
         Enabled8        =   -1  'True
         Position8       =   8
         Tip8            =   "Browse (Alt + Enter)"
         Visible8        =   0   'False
         Width8          =   16
         Enabled9        =   -1  'True
         Position9       =   9
         Tip9            =   " (Alt + Down, F4)"
         Visible9        =   -1  'True
         Width9          =   16
         Enabled10       =   -1  'True
         Position10      =   10
         Tip10           =   "Right Arrow (Alt + >)"
         Visible10       =   0   'False
         Width10         =   16
      End
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   9630
      Width           =   1305
   End
   Begin VB.TextBox Txt_Numorder 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁÌœ «·„Õ«”»Ì"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1035
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   11340
      Width           =   6465
      Begin MSDataListLib.DataCombo DcboDebitSide 
         Height          =   315
         Left            =   90
         TabIndex        =   30
         Top             =   270
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboCreditSide 
         Height          =   315
         Left            =   90
         TabIndex        =   32
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Index           =   12
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label LblDevID 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   270
         Width           =   1485
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·› —… :"
         Height          =   315
         Index           =   13
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÌœ:"
         Height          =   315
         Index           =   11
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› œ«∆‰"
         Height          =   285
         Index           =   10
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   885
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—› „œÌ‰"
         Height          =   285
         Index           =   9
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPTxtID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4320
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   645
      Left            =   16680
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox XPTxtVal 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   -1800
      Locked          =   -1  'True
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   7860
      Visible         =   0   'False
      Width           =   2145
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   9975
      _cx             =   17595
      _cy             =   1349
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
      Appearance      =   6
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "FrmExpenses40E.frx":11D5
      Caption         =   "   Œ·’ «Ê «” »⁄«œ«  «·«’Ê·  "
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1695
         TabIndex        =   6
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpenses40E.frx":1EAF
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   2
         Left            =   630
         TabIndex        =   7
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpenses40E.frx":2249
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   1
         Left            =   2220
         TabIndex        =   8
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpenses40E.frx":25E3
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   3
         Left            =   1155
         TabIndex        =   9
         Top             =   90
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpenses40E.frx":297D
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin MSAdodcLib.Adodc numbering 
         Height          =   585
         Left            =   4680
         Top             =   480
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc detect_no 
         Height          =   585
         Left            =   2640
         Top             =   600
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1032
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   " Õ—Ìﬂ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   3840
         Picture         =   "FrmExpenses40E.frx":2D17
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label LblShortcutKeys 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   510
         Width           =   5445
      End
   End
   Begin MSDataListLib.DataCombo XPCboExpensesType 
      Height          =   315
      Left            =   16080
      TabIndex        =   1
      Top             =   2760
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   7560
      TabIndex        =   13
      Top             =   9630
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   8580
      TabIndex        =   19
      Top             =   8940
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÃœÌœ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   7680
      TabIndex        =   20
      Top             =   8940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ⁄œÌ·"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   6750
      TabIndex        =   21
      Top             =   8940
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   3
      Left            =   5715
      TabIndex        =   22
      Top             =   8940
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " —«Ã⁄"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   4
      Left            =   4800
      TabIndex        =   23
      Top             =   8940
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ–›"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   24
      Top             =   8940
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdHelp 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   960
      TabIndex        =   25
      Top             =   8940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   5
      Left            =   3720
      TabIndex        =   26
      Top             =   8940
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   873
      ButtonStyle     =   1
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
      Height          =   2340
      Left            =   16080
      TabIndex        =   37
      Top             =   4440
      Visible         =   0   'False
      Width           =   10800
      _cx             =   19050
      _cy             =   4128
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorFixed  =   -2147483633
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExpenses40E.frx":697F
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      ExplorerBar     =   0
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
      Begin VB.PictureBox PicDes 
         BorderStyle     =   0  'None
         Height          =   1635
         Left            =   240
         RightToLeft     =   -1  'True
         ScaleHeight     =   1635
         ScaleWidth      =   2925
         TabIndex        =   40
         Top             =   960
         Visible         =   0   'False
         Width           =   2925
         Begin VB.TextBox TxtDes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   1125
            Left            =   30
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label LblDes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
            ForeColor       =   &H0000C8FF&
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   0
            Width           =   2445
         End
      End
      Begin VDSCOMBOLibCtl.SmartCombo CboDes 
         Height          =   315
         Left            =   240
         TabIndex        =   43
         ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
         Top             =   600
         Visible         =   0   'False
         Width           =   2955
         _cx             =   1973752924
         _cy             =   1973748268
         Alignment       =   0
         Appearance      =   3
         AutoSearch      =   0   'False
         BackColor       =   -2147483624
         BackgroundColor =   -2147483633
         BorderColor     =   0
         BorderVisible   =   -1  'True
         Caption         =   "SmartCombo1"
         CaptionAlignment=   4
         CaptionBackColor=   -2147483633
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionForeColor=   -2147483630
         CaptionHeight   =   15
         CaptionOnTop    =   0   'False
         CaptionMultiLine=   0
         Checkbox3D      =   0   'False
         CheckboxAlignment=   5
         CheckboxBackColor=   16777215
         CheckboxSize    =   13
         CheckboxValue   =   0
         BrowsePictureAlignment=   5
         BrowsePictureStretchH=   0
         BrowsePictureStretchV=   0
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
         ForeColor       =   0
         Gap             =   0
         HideSelection   =   -1  'True
         Locked          =   0   'False
         MaxLength       =   0
         MultiLine       =   0
         OnFocus         =   3
         PasswordChar    =   ""
         Picture         =   "FrmExpenses40E.frx":6AE5
         PictureAlignment=   5
         PictureBackColor=   -2147483624
         PictureStretchH =   0
         PictureStretchV =   0
         Redraw          =   -1  'True
         ScrollBar       =   0
         Style           =   0
         Text            =   ""
         UnderLine       =   0   'False
         Enabled0        =   -1  'True
         Position0       =   0
         Tip0            =   "Caption"
         Visible0        =   0   'False
         Width0          =   90
         Enabled1        =   -1  'True
         Position1       =   1
         Tip1            =   ""
         Visible1        =   -1  'True
         Width1          =   32
         Enabled2        =   -1  'True
         Position2       =   2
         Tip2            =   "Check Box (Space, Ctrl + Space)"
         Visible2        =   0   'False
         Width2          =   16
         Enabled3        =   -1  'True
         Position3       =   3
         Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
         Visible3        =   -1  'True
         Width3          =   145
         Enabled4        =   -1  'True
         Position4       =   4
         Tip4            =   "Left Spinner (Alt + Left)"
         Visible4        =   0   'False
         Width4          =   16
         Enabled5        =   -1  'True
         Position5       =   5
         Tip5            =   "Right Spinner (Alt + Right)"
         Visible5        =   0   'False
         Width5          =   16
         Enabled6        =   -1  'True
         Position6       =   6
         Tip6            =   "Up Spinner (Ctrl + Up)"
         Visible6        =   0   'False
         Width6          =   16
         Enabled7        =   -1  'True
         Position7       =   7
         Tip7            =   "Down Spinner (Ctrl + Down)"
         Visible7        =   0   'False
         Width7          =   16
         Enabled8        =   -1  'True
         Position8       =   8
         Tip8            =   "Browse (Alt + Enter)"
         Visible8        =   0   'False
         Width8          =   16
         Enabled9        =   -1  'True
         Position9       =   9
         Tip9            =   " (Alt + Down, F4)"
         Visible9        =   -1  'True
         Width9          =   16
         Enabled10       =   -1  'True
         Position10      =   10
         Tip10           =   "Right Arrow (Alt + >)"
         Visible10       =   0   'False
         Width10         =   16
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   255
      Left            =   9120
      TabIndex        =   38
      Top             =   11280
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "„—«ﬂ“ «· ﬂ·›…"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   192
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "FrmExpenses40E.frx":707F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   495
      Index           =   8
      Left            =   2760
      TabIndex        =   44
      Top             =   8940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   9
      Left            =   5640
      TabIndex        =   45
      Top             =   11400
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·‘Ìﬂ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ALLButtonS.ALLButton CmdRemove 
      Height          =   375
      Left            =   9600
      TabIndex        =   46
      Tag             =   "Delete Row"
      Top             =   10800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Õ–› ”ÿ—"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   0
      BCOLO           =   0
      FCOL            =   255
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmExpenses40E.frx":709B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      CausesValidation=   0   'False
      Height          =   375
      Index           =   10
      Left            =   4440
      TabIndex        =   67
      Top             =   9540
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
      Height          =   2340
      Left            =   120
      TabIndex        =   111
      Top             =   11160
      Width           =   10800
      _cx             =   19050
      _cy             =   4128
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   2
      GridLinesFixed  =   2
      GridLineWidth   =   2
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmExpenses40E.frx":70B7
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      ExplorerBar     =   0
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
   Begin ImpulseButton.ISButton ISButton1 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1920
      TabIndex        =   167
      Top             =   8940
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«·„—›ﬁ« "
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton CmdAttach 
      Height          =   375
      Left            =   3600
      TabIndex        =   168
      Top             =   9540
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "«·„—›ﬁ« "
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   390
      Index           =   8
      Left            =   8985
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   9645
      Width           =   900
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ"
      Height          =   255
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   9660
      Width           =   735
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Height          =   405
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   7680
      Width           =   6015
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9630
      Width           =   555
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   435
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9630
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "/"
      Height          =   435
      Index           =   6
      Left            =   1050
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   9630
      Width           =   165
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Index           =   7
      Left            =   1860
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   9630
      Width           =   1515
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«Ã„«·Ì"
      Height          =   285
      Index           =   2
      Left            =   -840
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   8040
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "·«„—"
      Height          =   285
      Index           =   5
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2520
      Width           =   1515
   End
End
Attribute VB_Name = "FrmExpenses40E"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim numbering_type As Integer
Dim departement_name  As String
Dim branch_no  As String
Dim RsNotes As ADODB.Recordset
Dim BolEditOnMainAccounts As Boolean
Dim Account_Code_dynamic3 As String
Dim Account_Code_dynamic4 As String
Dim group_id As Integer
Dim ProfitOrLose As Integer
Dim ProfitOrLoseValue As Double
Dim line_no As Integer
Dim LoseProfitValue As Double
Dim txtmyDes As String
Dim txtmyDesE As String
     
Function saveChequeBoxContents(NoteID As Double)

    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords

    If val(DCChequeBox.BoundText) = 0 Then Exit Function
 
 '   rs.Open "TblChecqueBoxContent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    StrSQL = "SELECT     * from dbo.TblChecqueBoxContent Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    rs.AddNew
    rs("noteid").value = NoteID
    rs("ChequeBoxID").value = val(DCChequeBox.BoundText)
            
    rs("RecordDate").value = XPDtbTrans.value
    rs("DueDate").value = DtpChequeDueDate.value
    rs("BankName").value = TXTBankName.text
    rs("ChequeNo").value = TxtChequeNumber.text
    rs("ChequeValue").value = val(XPTxtVal.text)
    
    rs("Remarks").value = DcboCreditSide.text
    rs("Deposited").value = 0
    rs("Collected").value = 0
    rs("CreditAccount").value = (DcboCreditSide.BoundText)
    rs.update
  
    rs.Close
End Function
                         
Function saveChequeBoxContents1(NoteID As Double)

    If SystemOptions.banks_Accounts3 = False Then Exit Function
    Dim i As Integer
    Dim rs As New ADODB.Recordset
 
    Dim StrSQL As String

    StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & NoteID
    Cn.Execute StrSQL, , adExecuteNoRecords
 
'    rs.Open "TblChecqueBoxContent1", Cn, adOpenStatic, adLockOptimistic, adCmdTable
   StrSQL = "SELECT     * from dbo.TblChecqueBoxContent1 Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
    If CboPayMentType.ListIndex = 1 Then
        rs.AddNew
        rs("noteid").value = NoteID
     
        rs("RecordDate").value = XPDtbTrans.value
        rs("DueDate").value = DtpChequeDueDate.value
        rs("BankID").value = val(DcboBankName.BoundText)
        rs("BankName").value = DcboBankName.text
        
        rs("ChequeNo").value = TxtChequeNumber.text
        rs("ChequeValue").value = val(XPTxtVal.text)
    
        rs("Remarks").value = Me.DcboDebitSide.text
        rs("Payed").value = 0
       
        rs("DepitAccount").value = (DcboDebitSide.BoundText)
        rs("notes_all").value = NoteID
      
        rs.update
    End If

    rs.Close
End Function

Private Sub ALLButton1_Click()
    On Error GoTo ErrTrap

    If DcCostCenter.BoundText <> "" Then

        MsgBox "·«Ì„ﬂ‰ «· Ê“Ì⁄ ⁄·Ï „—«ﬂ“ «· ﬂ·›… ·«‰ﬂ «Œ —   Ê“Ì⁄ ⁄«„ ⁄·Ï „—ﬂ“  ﬂ·›… „Õœœ", vbCritical
        Exit Sub
    End If

    Dim opr_id As Double

    If Not IsNumeric(Text1.text) Then Exit Sub
    'If Me.TxtModFlg.text = "N" Then
    opr_id = val(Me.Text1.text)
    'Else
    'opr_id = TxtDEV_NO.text
    'End If

    If Not Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) = "" Then
        If Not val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE"))) = 0 Then

            marakes_taklefa_tawze3.show
            
            marakes_taklefa_tawze3.value.Caption = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("VALUE")) ' Text4.Text
            marakes_taklefa_tawze3.depit_or_credit.Caption = "„œÌ‰"
            marakes_taklefa_tawze3.kedno = opr_id
            marakes_taklefa_tawze3.account_no = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode"))
            marakes_taklefa_tawze3.account_name = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountName"))
            marakes_taklefa_tawze3.lineno = Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
        
        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·«»œ „‰ «œŒ«· ﬁÌ„… ", vbCritical
            Else
                MsgBox "Enter Value First ", vbCritical
            End If

            Exit Sub
        End If
            
    End If

    marakes_taklefa_tawze3.opr_type = "”‰œ ’—›"
    marakes_taklefa_tawze3.opr_id = opr_id 'TxtDEV_NO.text 'Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo"))  'Text5.Text
    marakes_taklefa_tawze3.Adodc3.ConnectionString = connection_string
    marakes_taklefa_tawze3.Adodc3.CommandType = adCmdText
    marakes_taklefa_tawze3.Adodc3.RecordSource = "SELECT * FROM marakes_taklefa_temp  where kedno =" & opr_id & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1"))
    marakes_taklefa_tawze3.Adodc3.Refresh
    '    Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("distributed")) = "1"

    Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_AfterAutoCloseUp()
    PutData
    CboDes.Visible = False
End Sub

Private Sub CboType2_Change()

    Select Case CboType2.ListIndex

        Case 0
            CboPayMentType.Enabled = True
            TxtFASalesPrice.Enabled = True
            FraNote.Enabled = True

        Case 1, 2
 
            CboPayMentType.Enabled = False
            TxtFASalesPrice.Enabled = True
            FraNote.Enabled = False
    End Select
    CalculteValueAdded2

End Sub

Private Sub CboType2_Click()
    CboType2_Change
End Sub

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial1, "0612201401"

End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcChequeBox_Change()

    If DCChequeBox.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboCreditSide.BoundText = ModAccounts.GetMyAccountCodeRefined("TblBoxesData", "BoxID", val(Me.DCChequeBox.BoundText), "Account_Code1")
    End If

End Sub

Private Sub CboPayMentType_Change()

    If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        Me.DcboBox.text = ""
        DCVendor.text = ""
        DCAccounts1.text = ""
        DCChequeBox.text = ""

    End If

    If Me.CboPayMentType.ListIndex = 0 Then
        DCChequeBox.Enabled = False
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
    ElseIf Me.CboPayMentType.ListIndex = 1 Then

        If SystemOptions.ChequeBox = True Then
            TXTBankName.Visible = True
            DCChequeBox.Enabled = True
        Else
            TXTBankName.Visible = False
        End If

        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Me.DCVendor.Enabled = False
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
DCVendor.text = ""
DcboBox.text = ""
DCAccounts1.text = ""
        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
            lbl(19).Caption = " «—ÌŒ «·«” Õﬁ«ﬁ"
    
        Else
            lbl(18).Caption = "Cheque No"
            lbl(19).Caption = "Due Date"
        End If
    
    ElseIf Me.CboPayMentType.ListIndex = 2 Then '⁄„Ì·
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        DCChequeBox.Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DcboBox.Enabled = False
        Me.DCVendor.Enabled = True
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
    ElseIf Me.CboPayMentType.ListIndex = 3 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.text = ""
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
        Me.DCVendor.Enabled = False

 TXTBankName.Visible = False
        
        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·ÕÊ«·… "
            lbl(19).Caption = " «—ÌŒÂ«"
        Else
            lbl(18).Caption = "Transfer  No"
            lbl(19).Caption = "Date"
        End If
      
    ElseIf Me.CboPayMentType.ListIndex = 4 Then
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = False
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        DCVendor.BoundText = ""
        DcboBox.BoundText = ""
        DcboBankName.BoundText = ""
        DCAccounts1.Enabled = True
        DCChequeBox.Enabled = False
        '        DCAccounts1.text = ""
 
    ElseIf Me.CboPayMentType.ListIndex = 5 Then
        Me.lbl(16).Enabled = False
        Me.DcboBox.Enabled = False
        DcboBox.text = ""
        TXTBankName.Visible = False
        DCChequeBox.Enabled = False
        Me.lbl(19).Enabled = True
        Me.lbl(18).Enabled = True
        Me.lbl(17).Enabled = True
        Me.DcboBankName.Enabled = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        DCAccounts1.Enabled = False
        DCAccounts1.text = ""
        Me.DCVendor.Enabled = False

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(18).Caption = "—ﬁ„ «·‘Ìﬂ "
            lbl(19).Caption = " «—ÌŒÂ  "
        Else
            lbl(18).Caption = "Cheque No"
            lbl(19).Caption = "Date"
        End If
 
    Else
        Me.lbl(16).Enabled = True
        Me.DcboBox.Enabled = True
        Me.lbl(19).Enabled = False
        Me.lbl(18).Enabled = False
        Me.lbl(17).Enabled = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
        Me.DCVendor.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Function setfoxy()
    Text1.text = CStr(new_id("foxy", "id", "", True))

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id").value = Text1.text
 
    rs.update
    
End Function

Private Sub CboPaymentType1_Change()

    If Me.CboPaymentType1.ListIndex = 0 Then
        Fg_Journal.Visible = True
        VSFlexGrid1.Visible = False

    ElseIf Me.CboPaymentType1.ListIndex = 1 Then
        Fg_Journal.Visible = False
        VSFlexGrid1.Visible = True
    End If

End Sub

Private Sub CboPaymentType1_Click()
    CboPaymentType1_Change
End Sub

Private Sub CboType_Change()
    
    CboType_Click
    
End Sub

Private Sub CboType_Click()

    If Me.CboType.ListIndex = 1 Then ' Œ—Ìœ «’·
        TxtFASalesPrice.text = 0
        TxtFASalesPrice.Enabled = False
        CboPayMentType.Enabled = False
        FraNote.Enabled = False
        CboPayMentType.ListIndex = -1
        DCVendor.text = ""
        DcboBox.text = ""
        DcboBox.BoundText = 0
        DcboBankName.text = ""
        TxtChequeNumber.text = ""
        TxtFASalesPrice.Enabled = True
    Else
 
        TxtFASalesPrice.Enabled = True
        CboPayMentType.Enabled = True
        FraNote.Enabled = True
    
    End If

End Sub
Function GetVATBIll() As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     TOP 100 PERCENT dbo.DOUBLE_ENTREY_VOUCHERS.Vat"
sql = sql & " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
sql = sql & "                      dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
sql = sql & " Where (dbo.DOUBLE_ENTREY_VOUCHERS.FixedassetId = " & val(DcFixedAssets.BoundText) & ") And (dbo.Notes.notetype = 80)"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetVATBIll = IIf(IsNull(rs2("Vat").value), 0, rs2("Vat").value)
Else
GetVATBIll = 0
End If
End Function
Sub GetDataAssestVAT()
If Me.TxtModFlg.text <> "R" Then
    Dim astrSplitItems() As String
    Dim Result As String
    Dim YearMove As Double
    Dim YearNotMove As Double
    Dim MonthMove As Double
    Dim MonthNotMove As Double
   GetAssestMoveYearly DpPurchaseDate.value, YearMove, YearNotMove, MonthMove, MonthNotMove
   TxtYearMove.text = YearMove
   TxtYearNotMove.text = YearNotMove
   TxtMonthMove.text = MonthMove
   TxtMonthNotMove.text = MonthNotMove
    RetriveMoveassest
        Result = ExactAge(DpPurchaseDate.value, XPDtbTrans.value)
 If Result <> "" Then
    astrSplitItems = Split(Result, "-")
    TxtAgeYear.text = astrSplitItems(0)
    TxtAgeMonth.text = astrSplitItems(1)
  End If
  If RdMove(1).value = True Then
 TxtPeriodYear.text = val(Me.TxtYearNotMove) * 12 - (val(TxtAgeYear.text) * 12 + val(TxtAgeMonth.text))
 TxtPeriodMonth.text = 12 - val(Me.TxtAgeMonth.text)
 Else
 TxtPeriodYear.text = val(Me.TxtYearMove) * 12 - (val(TxtAgeYear.text) * 12 + val(TxtAgeMonth.text))
 TxtPeriodMonth.text = 12 - val(Me.TxtAgeMonth.text)
 End If
 TxtPeriodMonth.text = val(TxtPeriodYear.text) Mod 12
 TxtPeriodYear.text = val(TxtPeriodYear.text) \ 12
 
 TxtNetMonth.text = val(TxtPeriodMonth.text) + (val(TxtPeriodYear.text) * 12)
 TxtBillVAT.text = GetVATBIll()
 If RdMove(1).value = True Then
 If val(Me.TxtMonthNotMove.text) <> 0 Then
 TxtMonthVAT.text = Round(val(TxtBillVAT.text) / val(Me.TxtMonthNotMove.text), 2)
 Else
 TxtMonthVAT.text = 0
 End If
 Else
  If val(Me.TxtMonthMove.text) <> 0 Then
 TxtMonthVAT.text = Round(val(TxtBillVAT.text) / val(Me.TxtMonthMove.text), 2)
 Else
 TxtMonthVAT.text = 0
 End If
 End If
 TxtFATValue.text = val(TxtMonthVAT.text) * val(TxtNetMonth.text)
Dim account As String
PercentgValueAddedAccount_Transec XPDtbTrans.value, 12, 1, account
AccountVat.BoundText = account
   End If
End Sub

Private Sub Cmd_Click(Index As Integer)

    'On Error GoTo ErrTrap
    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
       
            TxtModFlg.text = "N"
            clear_all Me
            DcCostCenter.text = ""
            CboPaymentType1.ListIndex = 2
            ' XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            ' Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=3"))
        
            Me.DCboUserName.BoundText = user_id
            '        XPDtbTrans.SetFocus
          
            DtpChequeDueDate.value = Date
            setfoxy
            Me.dcBranch.BoundText = branch_id
            CuurentLogdata
GetDataAssestVAT
Dim Dcombos As New ClsDataCombos
Dcombos.GetFixedAssets Me.DcFixedAssets, , "0,1"


        Case 1
                            If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            Dim Msg  As String
                    
            '                  If SystemOptions.banks_Accounts3 = True Then
            '     If ChequeBoxOperations1(Val(Me.XPTxtID)) = False Then
            '         Msg = " ·« Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–… «·⁄„·Ì…"
            '         Msg = Msg & Chr(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
            '         MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '         Exit Sub
            '     End If
            ' End If
    
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
        
            If SystemOptions.ChequeBox = True And CboPayMentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ » ⁄œÌ· Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
    
            End If
      
            TxtModFlg.text = "E"
        
        Case 2
                                 If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify branch"
                Else
                    Msg = "Õœœ «·›—⁄"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                dcBranch.SetFocus
                Sendkeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            my_branch = Me.dcBranch.BoundText
    
            DcboBox_Change
            DcboBankName_Change
            DCVendor_Change
            DCAccounts1_Change
            DcChequeBox_Change
                   Dim AccountVATDept As String
If AccountVat.BoundText = "" And True = True And CheckAnyVAT(XPDtbTrans.value) = True Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
Else
MsgBox "Please Check the value-added settings"
End If
Exit Sub
End If
            SaveData
           
        Case 3
            Undo

        Case 4
                           If ChekClodePeriod(XPDtbTrans.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
        
            If SystemOptions.ChequeBox = True And CboPayMentType.ListIndex = 1 Then
         
                If ChequeBoxOperations(val(Me.XPTxtID)) = False Then
                    Msg = "·‰ Ì„ﬂ‰ «·”„«Õ »Õ–› Â–« «·⁄„·Ì…..!!!"
                    Msg = Msg & CHR(13) & "ÕÌÀ «‰Â« ”Ê› Ì‰ Ã ⁄‰Â« Œÿ« ›Ï   Õ«›Ÿ… «·‘Ìﬂ«  ·«‰Â  „ ⁄·ÌÂ« Õ—ﬂ«  «Ìœ«⁄ «Ê  Õ’Ì· "
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            Load FrmNotesSearch
            FrmNotesSearch.SearchType = 300
            FrmNotesSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            ViewDataList

        Case 8
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report (TxtSerial.text)

        Case 9
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_Cheque TxtChequeNumber.text, get_Cheque_report_no(val(DcboBankName.BoundText)), TxtSerial.text

        Case 10
    
            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            ShowGL_cc TxtSerial.text, , 200 ', val(TXT_A_NoteID.Text) 'NoteID
    
    End Select

    Exit Sub
ErrTrap:
End Sub

Function print_Cheque(Optional ChqueNum As String = "", Optional report_no As String = "", Optional serial As String)
    hide_logo = True
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From Expanses_Order  where ChqueNum='" & ChqueNum & "' and noteserial='" & TxtSerial & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    Else
        StrFileName = App.path & "\Reports\Chque\" & report_no & ".rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
    'MsgBox ToHijriDate(Date)

    xReport.ParameterFields(5).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 1, 2)
    xReport.ParameterFields(6).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 4, 2)
    xReport.ParameterFields(7).AddCurrentValue mId(ToHijriDate(DtpChequeDueDate.value), 9, 2)

    xReport.ParameterFields(8).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 1, 2)
    xReport.ParameterFields(9).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 4, 2)
    xReport.ParameterFields(10).AddCurrentValue mId(Format$(DtpChequeDueDate.value, "dd/mm/yyyy"), 9, 2)
    xReport.ParameterFields(11).AddCurrentValue CStr(txtto.text)
    xReport.ParameterFields(12).AddCurrentValue CStr(XPTxtVal.text)
    xReport.ParameterFields(13).AddCurrentValue CStr(Me.XPMTxtRemarks.text)
    xReport.ParameterFields(14).AddCurrentValue CStr(LblValue.Caption)
 
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

'Function print_report(Optional NoteSerial As String)
'
'    Dim MySQL As String
'    Dim RsData As New ADODB.Recordset
'    Dim xApp As New CRAXDRT.Application
'    Dim xReport As CRAXDRT.Report
''    Dim CViewer As ClsReportViewer
 '   Dim StrReportTitle As String
 '   Dim StrFileName As String
 '   Dim Msg As String
'
'    MySQL = "Select * From Expanses_Order  where noteserial='" & NoteSerial & "'"
'
'    'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'    '    MySQL = MySQL + " where RecordDate >=" & SQLDate(Me.DTPickerAccFrom.value, True) & ""
'    'End If
'    'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'    '    MySQL = MySQL + " and RecordDate <=" & SQLDate(Me.DTPickerAccTo.value, True) & ""
'    'End If
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
'    Else
'        StrFileName = App.path & "\Reports\" & "Expenses_order.rpt"
'    End If
'
'    If Dir(StrFileName) = "" Then
'        'GetMsgs 139, vbExclamation
'        Screen.MousePointer = vbDefault
''        Exit Function
 '   End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
''        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '       RsData.Close
 '       Set RsData = Nothing
 '       Screen.MousePointer = vbDefault
 '       Exit Function
 '   End If
'
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData
'
'    Dim cCompanyInfo As New ClsCompanyInfo
''
 '   If SystemOptions.UserInterface = ArabicInterface Then
 '       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
 '       ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
 ''       StrReportTitle = "" '& StrAccountName
  '  Else
 '
 '       xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
 '       'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
 '       xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
 '       StrReportTitle = ""
 '   End If
'
'    xReport.ParameterFields(3).AddCurrentValue user_name
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.title
'    xReport.ReportAuthor = App.title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, ""
'
'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault
'
'End Function


Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub
Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    SaveQRCode "notes_all", "NoteID", val(XPTxtID.text), TxtSerial1.text, (XPDtbTrans.value), _
        (txtNet.text), Picture2, 0, (txtVat2.text), (txtNet.text)
        
  MySQL = " SELECT     dbo.notes_all.NoteID,notes_all.Vat2,notes_all.Net, notes_all.QrCodeImage, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.NoteSerial, dbo.notes_all.Note_Value, dbo.notes_all.NoteSerial1, "
  MySQL = MySQL & "                    dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.notes_all.FAVType, dbo.notes_all.FAID,"
  MySQL = MySQL & "                    dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.notes_all.bill_type, dbo.notes_all.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
  MySQL = MySQL & "                    dbo.TblCustemers.Fullcode, dbo.notes_all.BoxID, TblBoxesData_1.BoxName, TblBoxesData_1.BoxNameE, dbo.notes_all.BankID, dbo.BanksData.BankName,"
  MySQL = MySQL & "                    dbo.BanksData.BankNamee, dbo.notes_all.ChqueNum, dbo.notes_all.ChequeBoxID, TblBoxesData_1.BoxName AS CheckBankName,"
  MySQL = MySQL & "                    TblBoxesData_1.BoxNameE AS CheckBankNameE, dbo.notes_all.DueDate, dbo.notes_all.Accountcode, dbo.ACCOUNTS.Account_Name,"
  MySQL = MySQL & "                     dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial, dbo.notes_all.general_des, dbo.notes_all.FASalesPrice, dbo.FixedAssets.code,"
  MySQL = MySQL & "                    dbo.notes_all.NewAssest,dbo.notes_all.NuminstallmCurr, dbo.notes_all.NuminstallmRemin, dbo.notes_all.NuminstallmExcu, dbo.notes_all.NuminstallmTotal,"
  MySQL = MySQL & "                    dbo.notes_all.CurrentValue, dbo.notes_all.AccDepre, dbo.notes_all.PurchasePrice, dbo.notes_all.LoseProfitValue, dbo.notes_all.NewNuminstallmRemin,"
  MySQL = MySQL & "                    dbo.notes_all.TxtExcludedValueInst, dbo.notes_all.ExcludedValueFixed, dbo.notes_all.ExcludedValuePrt, dbo.notes_all.ExcludedValue, dbo.notes_all.ExcludedType,"
  MySQL = MySQL & "                    dbo.notes_all.FATypeOp, dbo.notes_all.FAGroupID, dbo.FixedAssetsGroup.GroupName, dbo.FixedAssetsGroup.GroupNamee, dbo.FixedAssetsGroup.GroupCode,"
  MySQL = MySQL & "                    dbo.notes_all.NewCurrentValue, dbo.notes_all.BankName AS BankNameStr, dbo.notes_all.NoteCashingType, dbo.notes_all.NetMonth, dbo.notes_all.MonthVAT,"
  MySQL = MySQL & "                    dbo.notes_all.DpPurchaseDate, dbo.notes_all.AgeMonth, dbo.notes_all.AgeYear, dbo.notes_all.PeriodMonth, dbo.notes_all.PeriodYear, dbo.notes_all.BillVAT,"
  MySQL = MySQL & "                    dbo.notes_all.AsstMove, dbo.notes_all.MonthNotMove, dbo.notes_all.MonthMove, dbo.notes_all.YearNotMove, dbo.notes_all.YearMove,"
  MySQL = MySQL & "                    dbo.notes_all.AccountCodeVat , dbo.notes_all.TotalValue, dbo.notes_all.VATNO, dbo.notes_all.FATValue"
  MySQL = MySQL & " FROM         dbo.TblBoxesData TblBoxesData_2 RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.notes_all LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.FixedAssetsGroup ON dbo.notes_all.FAGroupID = dbo.FixedAssetsGroup.GroupID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.ACCOUNTS ON dbo.notes_all.Accountcode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBoxesData TblBoxesData_1 ON dbo.notes_all.ChequeBoxID = TblBoxesData_1.BoxID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.BanksData ON dbo.notes_all.BankID = dbo.BanksData.BankID ON TblBoxesData_2.BoxID = dbo.notes_all.BoxID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.FixedAssets ON dbo.notes_all.FAID = dbo.FixedAssets.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "  Where (dbo.notes_all.NoteType = 8028) And (dbo.notes_all.NoteID = " & val(XPTxtID.text) & ")"
 
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportsExpenses40E.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportsExpenses40E.rpt"
        End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
'xReport.ParameterFields(2).AddCurrentValue WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
    xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(IIf(IsNumeric(txtNet.text), txtNet.text, 0), "0.00"), 0, True, ".")
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
Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    Dim sql As String

    sql = "Delete  marakes_taklefa_temp where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
    Cn.Execute sql, , adExecuteNoRecords
    
    If CboPaymentType1.ListIndex = 0 Then
        If Fg_Journal.rows > 1 Then
            If Fg_Journal.rows = 2 Then
                Me.Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.Fg_Journal.rows > 1 Then
                    If Me.Fg_Journal.Row <> Me.Fg_Journal.FixedRows - 1 Then
                        Me.Fg_Journal.RemoveItem (Me.Fg_Journal.Row)
                    End If
                End If
            End If
        End If
            
        With Fg_Journal
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    ElseIf CboPaymentType1.ListIndex = 1 Then

        If VSFlexGrid1.rows > 1 Then
            If VSFlexGrid1.rows = 2 Then
                Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid1.rows > 1 Then
                    If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                        Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid1
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With
             
    ElseIf CboPaymentType1.ListIndex = 2 Then

        If VSFlexGrid2.rows > 1 Then
            If VSFlexGrid2.rows = 2 Then
                Me.VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
            Else

                If Me.VSFlexGrid2.rows > 1 Then
                    If Me.VSFlexGrid2.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                        Me.VSFlexGrid2.RemoveItem (Me.VSFlexGrid1.Row)
                    End If
                End If
            End If
        End If
            
        With VSFlexGrid2
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With
             
    Else
 
        Exit Sub
    End If

End Sub

Private Sub DCAccounts1_Change()

    If DCAccounts1.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = DCAccounts1.BoundText
    End If

End Sub

Private Sub DCAccounts1_Click(Area As Integer)
    DCAccounts1_Change
End Sub

Private Sub DCAccounts1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Unload Account_search
        Account_search.show
        Account_search.case_id = 194
            
    End If

End Sub

Private Sub DcboBankName_Change()

    'On Error Resume Next
    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code1")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
    
        If CboPayMentType.ListIndex = 3 Or CboPayMentType.ListIndex = 5 Then
                     
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
                    
        End If

        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

    End If

End Sub

Private Sub DcboBox_Change()

    If DcboBox.BoundText = "" Then Exit Sub
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    End If

End Sub

Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtSerial.text = ""
    TxtSerial1.text = ""
End Sub

Private Sub DcChequeBox_Click(Area As Integer)
    DcChequeBox_Change
End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 3
    End If

End Sub

Private Sub DcFixedAssets_Change()

    If val(DcFixedAssets.BoundText) = 0 Then Exit Sub
  
  
    DcFixedAssets_Click (0)
End Sub

Private Sub DcFixedAssets_Click(Area As Integer)
    Dim AccDepreciation As Double
    Dim RemianInstallments As Double
    Dim CurrentInstalmentNo As Double
    Dim Installmentvalue As Double
    Dim NewAccDepreciation As Double
    Dim FixedAsssetid As Integer
    Dim purchaseprice As Double
    Dim FixedAssetName As String
    Dim fullcode As String
    Dim KhordaPrice As Double
    Dim PurchaseDate As Date
    Dim branch_no As Integer
    If val(DcFixedAssets.BoundText) = 0 Then Exit Sub
    FixedAsssetid = val(DcFixedAssets.BoundText)
    Me.TxtFASalesPrice = 0

    GetFixedAssetHistory FixedAsssetid, AccDepreciation, RemianInstallments, CurrentInstalmentNo, Installmentvalue, NewAccDepreciation, purchaseprice, FixedAssetName, , fullcode, KhordaPrice, group_id, , , PurchaseDate, branch_no
 dcBranch.BoundText = branch_no
    TxtPurchasePrice.text = purchaseprice
    TxtAccDepre.text = AccDepreciation
    ' TxtCurrentValue = TxtPurchasePrice.text - (TxtAccDepre.text + KhordaPrice)
 Me.txtAdditions.text = GetAddValue(val(DcFixedAssets.BoundText))


    txtCurrentValue = TxtPurchasePrice.text - TxtAccDepre.text + GetAddValue(val(DcFixedAssets.BoundText))
 '''''''''''''''''''''''''''''''
 txtCountF = 1
 TxtNewCurrentValue = txtCurrentValue
 
  '  TxtCurrentValue.text = val(TxtCurrentValue.text)
  '  TxtFixeCurValue.text = TxtCurrentValue.text
 TxtNuminstallmCurr.text = Installmentvalue
 TxtNuminstallmRemin.text = RemianInstallments + GetQstAddNo(val(DcFixedAssets.BoundText))
 TxtNuminstallmExcu.text = CurrentInstalmentNo
 TxtNuminstallmTotal.text = RemianInstallments + CurrentInstalmentNo + GetQstAddNo(val(DcFixedAssets.BoundText))
DpPurchaseDate.value = PurchaseDate
 '''''
    TxtFASalesPrice_Change
    WriteDev
GetDataAssestVAT
End Sub
Function GetQstAddNo(Optional FixedID As Integer = 0) As Double
If FixedID <> 0 Then
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = sql & " SELECT     SUM(QstIncNo) AS SmQst, FixedID"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & "  Where (FixedID = " & FixedID & ")"
sql = sql & " GROUP BY FixedID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GetQstAddNo = IIf(IsNull(Rs8("SmQst").value), 0, Rs8("SmQst").value)
Else
GetQstAddNo = 0
End If
End If
End Function
Function GetAddValue(Optional FixedID As Integer = 0) As Double
If FixedID <> 0 Then
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
sql = sql & " SELECT     SUM(AddValue) AS SmAddValue, FixedID"
sql = sql & " From dbo.TblAdditionsAssest"
sql = sql & "  Where (FixedID = " & FixedID & ")"
sql = sql & " GROUP BY FixedID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GetAddValue = IIf(IsNull(Rs8("SmAddValue").value), 0, Rs8("SmAddValue").value)
Else
GetAddValue = 0
End If
End If
End Function

Function WriteDev()

    If 1 = 1 Then

        If SystemOptions.AssetAccount1 = True Then
            If val(TxtFASalesPrice.text) > val(TxtNewCurrentValue.text) Then
                DCAccounts.BoundText = get_FixedAsset_Account(group_id, branch_id, "Account_Code3")
            Else
                DCAccounts.BoundText = get_FixedAsset_Account(group_id, branch_id, "Account_Code4")
            End If
                             
        Else
              
            If val(TxtFASalesPrice.text) > val(TxtNewCurrentValue.text) Then
                Account_Code_dynamic3 = get_account_code_branch(66, my_branch)
                                            
                If Account_Code_dynamic3 = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic3 = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ     Õ”«» «—»«Õ »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
                                             
                    End If
                End If

                DCAccounts.BoundText = Account_Code_dynamic3
            Else
                Account_Code_dynamic4 = get_account_code_branch(67, my_branch)
                                            
                If Account_Code_dynamic4 = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic4 = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ  Õ”«» Œ”«—… »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
                                             
                    End If
                End If

                DCAccounts.BoundText = Account_Code_dynamic4
            End If
              
        End If
          
    End If

ErrTrap:
End Function

Private Sub DcFixedAssets_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
    'Wael
        FixedAssetsSearch.RetrunType = 12
        FixedAssetsSearch.show vbModal
  
    End If

End Sub

Private Sub DCVendor_Change()

    If DCVendor.BoundText = "" Then Exit Sub

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Text2.text = Me.DCVendor.BoundText
End Sub

Private Sub DCVendor_Click(Area As Integer)
    DCVendor_Change
End Sub

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)

            Case "ExpensesID"
              
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If

            Case "value", "opr_fullcode"
                Dim sgl As String
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")
                
                If checkitems(project_id, .TextMatrix(Row, .ColIndex("opr_fullcode")), val(.TextMatrix(Row, .ColIndex("Value")))) = False Then
                    .TextMatrix(Row, .ColIndex("Value")) = 0
                End If
               
                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Function calcnets()

    If Me.CboPaymentType1.ListIndex = 0 Then

        With Fg_Journal
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    ElseIf Me.CboPaymentType1.ListIndex = 1 Then

        With Me.VSFlexGrid1
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    ElseIf Me.CboPaymentType1.ListIndex = 2 Then

        With Me.VSFlexGrid2
            Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        End With

    End If

End Function

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
                '  Cancel = True
            
            Case "Order_No"
                .ComboList = ""
        End Select

    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.cell(flexcpData, r, c)) <> "String" Then
            TxtDes.text = ""
        Else
            '
            TxtDes.text = Fg_Journal.cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                    Order_no_search.show
                    Order_no_search.RetrunType = 4
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.show
                    FrmExpensesSearch.RetrunType = 2
                End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)

            Case "AccountName"
                StrSQL = "select * from Expenses_accounts"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                  
            Case "opr_fullcode"
                Dim project_id As Integer
                project_id = get_project_id(dcproject.BoundText, "expanses_account")

                If SystemOptions.Items_or_operation = 1 Then
                    StrSQL = "  select fullcode,name from terms_operations where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,name", "fullcode")
                ElseIf SystemOptions.Items_or_operation = 0 Then
                    StrSQL = "  select fullcode,des from projects_des where project_id=" & project_id
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                    StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode,des", "fullcode")
         
                End If

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String

    On Error GoTo ErrTrap

    StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
    fill_combo Me.DcCostCenter, StrSQL
    
    ScreenNameArabic = " Œ·’ «Ê «” »⁄«œ«  «·«’Ê·"
    ScreenNameEnglish = "Disposal of assets"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("FillData").Picture
    Resize_Form Me
    AddTip
    SetDtpickerDate XPDtbTrans
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetExpensesType XPCboExpensesType
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.XPCboExpensesType
    Dcombos.GetFixedAssetsGroup DcbGroup
    Dcombos.GetAccountingCodes Me.DcboDebitSide
    Dcombos.GetAccountingCodes Me.DcboCreditSide
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetAccountingCodes Me.DCAccounts, True
    Dcombos.GetAccountingCodes Me.DCAccounts1, True
    Dcombos.GetFixedAssets Me.DcFixedAssets
    Dcombos.GetAccountingCodes AccountVat
    Dcombos.GetChequeBox Me.DCChequeBox

    With Me.CboPayMentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "‘Ìﬂ"
        .AddItem "«Ã·"
        .AddItem "ÕÊ«·…"
        .AddItem "Õ”«»"
        .AddItem "‘Ìﬂ „Õ’·"
       
    End With

    With Me.CboPaymentType1
        .Clear
        .AddItem "„’«—Ì›"
        .AddItem "Õ”«»« "
        .AddItem "‘—«¡ «’· À«» "
    End With

    With Me.CboType
        .Clear
        .AddItem " Œ·’ „‰ «’· »«·ﬂ«„·"
        .AddItem "«” »⁄«œ „‰ «’·"
    
    End With

    With Me.CboType2
        .Clear
        .AddItem "»«·»Ì⁄"
        .AddItem "»«· Œ—Ìœ"
        .AddItem "«’· „‰›’·"
    End With

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    StrSQL = " select expanses_account,Project_name from projects  where not(expanses_account is null)"
    fill_combo dcproject, StrSQL

    'StrSQL = " select  CusID, CusName from TblCustemers  where Type=3"
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " Select CusID,CusName From TblCustemers Where Type=1 or CustomerandVendor=1"
    Else
        StrSQL = " Select CusID,CusNamee From TblCustemers Where Type=1 or CustomerandVendor=1"
    End If

    fill_combo Me.DCVendor, StrSQL

    Set rs = New ADODB.Recordset
    StrSQL = "select * From notes_all where notetype=8028  "
    StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
    
         If SystemOptions.FixedCustomer = 1 Then
                              StrSQL = StrSQL & " and  UserID = " & user_id
                               End If
                               
              If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    End If
    
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    hide_logo = False
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
        Set rs = Nothing
    End If

    Set TTP = Nothing
    'Set EmpReport = Nothing
    TTD.Destroy
    Exit Sub
ErrTrap:
End Sub

Private Sub CboDes_ButtonClick(ByVal ButtonID As VDSCOMBOLibCtl.vdsButtonID, _
                               ByVal SpinningEnded As Boolean)

    If ButtonID = vdsDownArrow Then
        If CboDes.IsDropped = False Then
            If PicHeight > 0 Then
                PicDes.Height = PicHeight
                PicDes.Width = PicWidth
            Else
                PicDes.Width = CboDes.Width - 10
                PicDes.Height = CboDes.Height * 8
            End If

            Debug.Print PicHeight
            Debug.Print PicWidth
            TxtDes.Visible = True
            TxtDes.text = Fg_Journal.cell(flexcpData, Fg_Journal.Row, Fg_Journal.ColIndex("Des"))
            CboDes.DropDown PicDes.hWnd, vdsRightToLeft, vdsBottomToDown, vdsDownArrow, True, vdsSoftResize
            Debug.Print PicDes.Height & "Pic H " & "-----" & PicDes.Width & "Pic W"
        Else
            CboDes.CloseUp
        End If
    End If

End Sub

Private Sub CboDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Sendkeys "{F4}"
    End If

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub PicDes_Resize()

    With PicDes
        LblDes.Move .ScaleLeft, .ScaleTop, .ScaleWidth, LblDes.Height
        TxtDes.Move .ScaleLeft, .ScaleTop + LblDes.Height, .ScaleWidth, .ScaleHeight - LblDes.Height
        '    PicHeight = PicDes.Height
        '    PicWidth = PicDes.Width
    End With

End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search.show
        Order_no_search.RetrunType = 1
    End If

End Sub

Private Sub txtCountF_Change()
If (TxtModFlg.text = "R" Or TxtModFlg.text = "") Then Exit Sub
    If val(txtCountF) <> 0 Then
        TxtNewCurrentValue = val(txtCurrentValue) / val(txtCountF)
    End If
End Sub

Private Sub TxtDes_LostFocus()
    PicHeight = PicDes.Height
    PicWidth = PicDes.Width
    CboDes.CloseUp
    CboDes.Visible = False
End Sub

Private Sub TxtDes_KeyDown(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyEscape Then
        PutData
        CboDes.CloseUp
    End If

End Sub

Private Sub TxtExcludedValue_Change()
If RdExcluded(1).value = True Then
    TxtExcludedValuePrt = val(txtCurrentValue) * val(TxtExcludedValue) / 100
Else
    TxtExcludedValuePrt = val(TxtExcludedValue)
End If
 'txtCurrentValue = TxtPurchasePrice.text - TxtAccDepre.text + GetAddValue(val(DcFixedAssets.BoundText))
TxtLoseProfitValue = TxtExcludedValuePrt

If val(TxtNuminstallmRemin) <> 0 Then
    TxtNuminstallmCurr = Round(val(txtCurrentValue) / val(TxtNuminstallmRemin), 2)
End If
'TxtNuminstallmRemin
'TxtCurrentValue
End Sub

Private Sub TxtExcludedValuePrt_Change()
If (TxtModFlg.text = "R" Or TxtModFlg.text = "") Then Exit Sub
    Dim sql  As String
    sql = "Select Sum(netvalue) netvalue from TblFixedAssestTmpValue Where TransId = " & val(TxtSerial1)
    sql = sql & " and FixedID = " & val(DcFixedAssets.BoundText)
    Dim rsDummyFi As New ADODB.Recordset
    rsDummyFi.Open sql, Cn, adOpenKeyset, adLockOptimistic
    Dim mNoteValue As Double
    If Not rsDummyFi.EOF Then
        mNoteValue = val(rsDummyFi!netvalue & "")
        
    Else
'        mNoteValue = val(TxtExcludedValueFixed)
    End If
    rsDummyFi.Close
    sql = "Select CurrentValue,(FixedAssets.Quantity) as Quantity from FixedAssets   Where id = " & val(DcFixedAssets.BoundText)
     rsDummyFi.Open sql, Cn, adOpenKeyset, adLockOptimistic
    If Not rsDummyFi.EOF Then
        If val(rsDummyFi!Quantity & "") <> 0 Then
            'mNoteValue = val((rsDummyFi!currentvalue & "") * val(rsDummyFi!Quantity & "")) + mNoteValue
            mNoteValue = val((rsDummyFi!currentvalue & "")) + mNoteValue
        Else
            mNoteValue = (rsDummyFi!currentvalue & "") + mNoteValue
        End If
    End If
    txtCurrentValue = mNoteValue
    
    txtCurrentValue = TxtPurchasePrice.text - TxtAccDepre.text + GetAddValue(val(DcFixedAssets.BoundText))
TxtExcludedValueFixed = val(txtCurrentValue) - val(TxtExcludedValuePrt)
TxtNewCurrentValue = TxtExcludedValueFixed
txtCurrentValue = TxtNewCurrentValue

TxtLoseProfitValue = TxtExcludedValuePrt
If val(TxtNuminstallmRemin) <> 0 Then
    
End If
End Sub

Public Sub TxtFASalesPrice_Change()
 
    If (TxtModFlg.text = "R" Or TxtModFlg.text = "") Then Exit Sub
    'DcFixedAssets_Click (0)

    LoseProfitValue = val(TxtFASalesPrice) - val(TxtNewCurrentValue)
    TxtLoseProfitValue.text = Round(Abs(LoseProfitValue), 2)

    If LoseProfitValue > 0 Then
        TxtLoseProfitValue.ForeColor = vbGreen
    ElseIf LoseProfitValue < 0 Then
        TxtLoseProfitValue.ForeColor = vbRed
    Else
        TxtLoseProfitValue.ForeColor = vbBlack
    End If
    
    CalculteValueAdded2
    WriteDev
End Sub
Public Sub CalculteValueAdded2(Optional posDelete As Boolean = False)


If CboType2.ListIndex <> 0 Then txtVat2 = "": txtNet = TxtFASalesPrice:  Exit Sub
If val(txtNet) = 0 Then txtNet = val(TxtFASalesPrice) + val(txtVat2)
'If SystemOptions.PriceWithVAT = True Then Exit Sub
If (TxtModFlg.text = "R" Or TxtModFlg.text = "") Then Exit Sub
 Dim Percentg As Double
Dim AccountVATCreit As String
Dim cCompanyInfo As New ClsCompanyInfo
If True = True Then
'If TransType = 9 And ReturnSales = True Then

  
    
    If SystemOptions.AllItemInVAT = True Then
        Percentg = val(cCompanyInfo.VATItems)
    Else
      PercentgValueAddedAccount_Transec XPDtbTrans.value, 12, 1, AccountVATCreit, Percentg
        
    End If
'    If Percentg = -1 Then
'        Percentg = 0
'    Else
'
'    End If
   
'  If CboType.ListIndex = 1 Then
'
'    TxtVAt2 = val(TxtNewCurrentValue) * Percentg / 100
'
'
'     txtNet = val(TxtNewCurrentValue) + val(TxtVAt2)
'
'    Else
     txtVat2 = val(TxtFASalesPrice) * Percentg / 100
     
     
     txtNet = val(TxtFASalesPrice) + val(txtVat2)

'    End If
 
   
    

End If

End Sub

Private Sub TxtModFlg_Change()

    'On Error GoTo ErrTrap
    Select Case Me.TxtModFlg.text

        Case "R"
        Dim Dcombos As New ClsDataCombos
Dcombos.GetFixedAssets Me.DcFixedAssets
            DcFixedAssets.Enabled = False
        
            Me.VSFlexGrid1.Enabled = False
            Me.Fg_Journal.Enabled = False
            Frame1.Enabled = False
        
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            CmdRemove.Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            XPTxtVal.locked = True
            '        XPCboProfLevel.Locked = True
            '        XPTxtProfMail.Locked = True
            '        XPTxtPhone.Locked = True
            '        XPTxtMobile.Locked = True
            XPMTxtRemarks.locked = True
            XPCboExpensesType.locked = True
            Me.DcboBox.locked = True
            XPDtbTrans.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            
            End If

        Case "N"
        
            DcFixedAssets.Enabled = True
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True

            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            CmdRemove.Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
            XPDtbTrans.value = Date

        Case "E"
 
            DcFixedAssets.Enabled = False
            Me.VSFlexGrid1.Enabled = True
            Me.Fg_Journal.Enabled = True
            Frame1.Enabled = True
       
            CmdRemove.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            XPTxtVal.locked = False
            '        XPCboProfLevel.Locked = False
            '        XPTxtProfMail.Locked = False
            '        XPTxtPhone.Locked = False
            '        XPTxtMobile.Locked = False
            XPMTxtRemarks.locked = False
            XPCboExpensesType.locked = False
            Me.DcboBox.locked = False
            XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Sub RetriveMoveassest()
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " select  * from FixedAssetsGroup where GroupID in (select group_id from FixedAssets where ID =" & val(DcFixedAssets.BoundText) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
    If Not IsNull(rs2("AsstMove").value) Then
    If (rs2("AsstMove").value) = 1 Then
    RdMove(1).value = True
    Else
    RdMove(0).value = True
    End If
    Else
    RdMove(0).value = True
    End If
End If
End Sub

Private Sub TxtNewCurrentValue_Change()
If CboType.ListIndex = 1 Then
    Dim LoseProfitValue  As Double
    If val(txtCountF) = 0 Then
        LoseProfitValue = val(TxtFASalesPrice) - val(TxtNewCurrentValue)  'val(TxtNewCurrentValue) - val(TxtCurrentValue)
    Else
        'LoseProfitValue = val(TxtNewCurrentValue) - (val(TxtCurrentValue) / val(txtCountF))
        LoseProfitValue = val(TxtFASalesPrice) - val(TxtNewCurrentValue)
    End If
    TxtLoseProfitValue.text = Round(Abs(LoseProfitValue), 2)
    If LoseProfitValue > 0 Then
        TxtLoseProfitValue.ForeColor = vbGreen
    ElseIf LoseProfitValue < 0 Then
        TxtLoseProfitValue.ForeColor = vbRed
    Else
        TxtLoseProfitValue.ForeColor = vbBlack
    End If
    
    CalculteValueAdded2
End If
WriteDev
End Sub

Public Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, _
                                 ByVal Col As Long)
    'check_cost_center
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
 
    With VSFlexGrid1

        Select Case .ColKey(Col)
    
            Case "Value"
                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))

            Case "DebitValue", "CreditValue"

                'remove destribution
     
                ' sgl = "update  marakes_taklefa_temp  set value=0 where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                ' Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))
            
                If .ColKey(Col) = "DebitValue" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0
                    ' Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                 
                    '    Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValue" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0
                    ' Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    ' Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '     Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '       Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If

                .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
            
            Case "DebitValueE", "CreditValueE"
                .TextMatrix(Row, Col) = val(.TextMatrix(Row, Col))

                If .ColKey(Col) = "DebitValueE" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignRightCenter
                    .TextMatrix(Row, .ColIndex("CreditValueE")) = 0
                    .TextMatrix(Row, .ColIndex("CreditValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("DebitValue")) = .TextMatrix(Row, .ColIndex("DebitValueE"))
                    End If

                    '
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                ElseIf .ColKey(Col) = "CreditValueE" Then
                    .cell(flexcpAlignment, Row, .ColIndex("AccountName")) = flexAlignLeftCenter
                    .TextMatrix(Row, .ColIndex("DebitValueE")) = 0
                    .TextMatrix(Row, .ColIndex("DebitValue")) = 0

                    If .TextMatrix(Row, .ColIndex("rate")) <> "" Then
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE")) * .TextMatrix(Row, .ColIndex("rate"))
                    Else
                        .TextMatrix(Row, .ColIndex("CreditValue")) = .TextMatrix(Row, .ColIndex("CreditValueE"))
                    End If
                 
                    '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), .Rows - 1, .ColIndex("DebitValue"))
                    '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), .Rows - 1, .ColIndex("CreditValue"))
                    '      Me.TxtTotalDebit.text = Format(Me.TxtTotalDebit.text, SystemOptions.SysDefCurrencyForamt)
                    '        Me.TxtTotalCredit.text = Format(Me.TxtTotalCredit.text, SystemOptions.SysDefCurrencyForamt)
                       
                End If
            
            Case "Account_Serial"
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                StrSQL = "SELECT ACCOUNTS.cost_center, ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Serial='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    If BolEditOnMainAccounts = False Then
                        'If LastAccount(rs("Account_Code").value) = False Then
                        '    .TextMatrix(Row, Col) = ""
                        '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                        '    Exit Sub
                        'End If
                    End If

                    .TextMatrix(Row, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                    Else
                        .TextMatrix(Row, .ColIndex("AccountName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                    
                    End If
                    
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), 0, rs("cost_center").value)
                    
                    Dim rs2 As ADODB.Recordset
                    Dim My_SQL As String

                    If IsNull(rs("currenct_code").value) Then

                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                    
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo xx
                    End If

                    My_SQL = "  select * from currency WHERE id=" & val(rs("currenct_code").value)

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), 1, rs2.Fields("rate").value)
xx:
                Else
                    GetMsgs 130, vbExclamation
                    .TextMatrix(Row, Col) = ""
                    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    Exit Sub
                End If

                rs.Close
                Set rs = Nothing

            Case "AccountName"
        
                'sgl = "Delete  marakes_taklefa_temp  where kedno =" & Val(Text1.text) & " and account_no='" & Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("AccountCode")) & "' and  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                'Cn.Execute sgl, , adExecuteNoRecords
    
                .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)

                If LngRow <> -1 Then
                    'Msg = "Â–« «·Õ”«» „ÊÃÊœ „”»ﬁ«  ›Ï «·”ÿ— " & .TextMatrix(LngRow, .ColIndex("LineNo"))
                    'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    '.TextMatrix(Row, Col) = ""
                    '.TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Exit Sub
                End If

                Set ClsAcc = New ClsAccounts

                If BolEditOnMainAccounts = False Then
                    'If LastAccount(StrAccountCode) = False Then
                    '    .TextMatrix(Row, Col) = ""
                    '    .TextMatrix(Row, .ColIndex("AccountCode")) = ""
                    'Else

                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                    'End If
                Else
                    .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
 
                    .TextMatrix(Row, .ColIndex("Account_Serial")) = ClsAcc.Get_Account_Serial(StrAccountCode)
                End If

                Set ClsAcc = Nothing
            
                StrSQL = "SELECT ACCOUNTS.cost_center ,ACCOUNTS.currenct_code,ACCOUNTS.rate, ACCOUNTS.Account_Code, ACCOUNTS.Account_Name," & "ACCOUNTS.Parent_Account_Code, ACCOUNTS.last_account," & "ACCOUNTS.Account_NameEng,ACCOUNTS.Account_Serial" & " From ACCOUNTS Where ACCOUNTS.Account_Name='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    .TextMatrix(Row, .ColIndex("cost_center")) = IIf(IsNull(rs("cost_center").value), vbFalse, rs("cost_center").value)
            
                    'Dim rs2 As ADODB.Recordset
                    'Dim My_SQL As String
                    If IsNull(rs("currenct_code").value) Then
                        .TextMatrix(Row, .ColIndex("currenct_code")) = ""
                        .TextMatrix(Row, .ColIndex("rate")) = "1"
                    
                        GoTo ll
                    End If

                    My_SQL = "  select * from currency WHERE id=" & rs("currenct_code").value

                    Set rs2 = New ADODB.Recordset
                    rs2.CursorLocation = adUseClient
                    rs2.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    .TextMatrix(Row, .ColIndex("currenct_code")) = IIf(IsNull(rs2.Fields("code").value), "", rs2.Fields("code").value)
                    
                    .TextMatrix(Row, .ColIndex("rate")) = IIf(IsNull(rs2.Fields("rate").value), "", rs2.Fields("rate").value)
ll:
                End If

        End Select

        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ReLineGrid

    End With

End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid1

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "Value"
                .ComboList = ""

            Case "Account_Serial"
                .ComboList = ""
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub VSFlexGrid1_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 80

    End If

End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid1

        Select Case .ColKey(Col)

            Case "AccountName"
                
                'Full Path Display
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_NameEng As FirstName," & "ACCOUNTS_1.Account_NameEng As ParentName, ACCOUNTS_2.Account_NameEng As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '   If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '   End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_NameEng"
                    End If
                
                Else
                
                    StrSQL = "SELECT ACCOUNTS.Account_Code, ACCOUNTS.Account_Name As FirstName," & "ACCOUNTS_1.Account_Name As ParentName, ACCOUNTS_2.Account_Name As RootName " & " FROM (ACCOUNTS INNER JOIN ACCOUNTS AS ACCOUNTS_1 ON " & "ACCOUNTS.Parent_Account_Code = ACCOUNTS_1.Account_Code) " & "INNER JOIN ACCOUNTS AS ACCOUNTS_2 ON ACCOUNTS_1.Parent_Account_Code" & "= ACCOUNTS_2.Account_Code Where ACCOUNTS.Account_Code <>'r' "

                    '     If ChkLastAccount.value = vbChecked Then
                    If SystemOptions.SysDataBaseType = AccessDataBase Then
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account= True) "
                    Else
                        StrSQL = StrSQL + " And(ACCOUNTS.last_account=1)"
                    End If

                    '     End If
                    If OptSort(1).value = True Then
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Code"
                    Else
                        StrSQL = StrSQL + " Order By ACCOUNTS.Account_Name"
                    End If
                
                End If
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "RootName,ParentName,*FirstName", "Account_Code")
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub VSFlexGrid2_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid2

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                Dim GroupID As Integer
                Dim branch_id As Integer
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
            
                StrSQL = "select * from FixedAssets where id=" & val(StrAccountCode)
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    GroupID = IIf(IsNull(rs("group_id").value), "", rs("group_id").value)
                    .TextMatrix(Row, .ColIndex("groupid")) = GroupID
                    branch_id = IIf(IsNull(rs("Branch_NO").value), "", rs("Branch_NO").value)
                    .TextMatrix(Row, .ColIndex("branch_id")) = branch_id
              
                Else
                    .TextMatrix(Row, .ColIndex("groupid")) = 0
                    GroupID = 0
                    branch_id = 0
                    .TextMatrix(Row, .ColIndex("branch_id")) = 0
                End If
              
                .TextMatrix(Row, .ColIndex("AccountCode")) = get_FixedAsset_Account(GroupID, branch_id)
               
                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid

End Sub

Private Sub VSFlexGrid2_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid2
 
        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
    
        End Select

    End With

End Sub

Private Sub VSFlexGrid2_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid2

        Select Case .ColKey(Col)

            Case "AccountName"
                StrSQL = "select * from FixedAssets where New_or_opening=0 and PurchasePrice=0 order by Name"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = Fg_Journal.BuildComboList(rs, "Name", "Id")

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
         
        End Select

    End With

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    Select Case Index

        Case 0

            If Not (rs.EOF Or rs.BOF) Then
                rs.MovePrevious

                If rs.BOF Then rs.MoveFirst
            End If

        Case 1

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveFirst
            End If

        Case 2

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If

        Case 3

            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveNext

                If rs.EOF Then rs.MoveLast
            End If

    End Select

    Retrive
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As String = "")
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid = "-1" Then
        Cmd_Click (0)
    End If

    If Lngid <> "" Then
        '  If XPTxtID.text <> 0 Then
        rs.Find "NoteID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            clear_all Me

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "›« Ê—… €Ì— „”Ã·… ", vbInformation
            Else
                MsgBox " Un Refistered Bill ", vbInformation
            End If

            Exit Sub
        End If

        '  End If
    End If

    If Not IsNull(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If
''//////////////////////
Me.TxtYearMove.text = IIf(IsNull(rs("YearMove").value), "", rs("YearMove").value)
Me.TxtYearNotMove.text = IIf(IsNull(rs("YearNotMove").value), "", rs("YearNotMove").value)
Me.TxtMonthMove.text = IIf(IsNull(rs("MonthMove").value), "", rs("MonthMove").value)
Me.TxtMonthNotMove.text = IIf(IsNull(rs("MonthNotMove").value), "", rs("MonthNotMove").value)
Me.TxtFATValue.text = IIf(IsNull(rs("FATValue").value), "", rs("FATValue").value)
Me.TxtBillVAT.text = IIf(IsNull(rs("BillVAT").value), "", rs("BillVAT").value)
Me.TxtPeriodYear.text = IIf(IsNull(rs("PeriodYear").value), "", rs("PeriodYear").value)
Me.TxtPeriodMonth.text = IIf(IsNull(rs("PeriodMonth").value), "", rs("PeriodMonth").value)
Me.TxtAgeYear.text = IIf(IsNull(rs("AgeYear").value), "", rs("AgeYear").value)
Me.TxtAgeMonth.text = IIf(IsNull(rs("AgeMonth").value), "", rs("AgeMonth").value)
Me.AccountVat.BoundText = IIf(IsNull(rs("AccountCodeVat").value), "", rs("AccountCodeVat").value)
DpPurchaseDate.value = IIf(IsNull(rs("DpPurchaseDate").value), Date, rs("DpPurchaseDate").value)
Me.TxtMonthVAT.text = IIf(IsNull(rs("MonthVAT").value), "", rs("MonthVAT").value)
Me.TxtNetMonth.text = IIf(IsNull(rs("NetMonth").value), "", rs("NetMonth").value)
        txtVat2 = IIf(IsNull(rs("Vat2").value), "", rs("Vat2").value)
    txtNet = IIf(IsNull(rs("Net").value), "", rs("Net").value)
If Trim(txtNet) = "" Then
    txtNet = val(TxtFASalesPrice) + val(txtVat2)
End If
If Not IsNull(rs("AsstMove").value) Then
If rs("AsstMove").value = 1 Then
RdMove(1).value = True
Else
RdMove(0).value = True
End If
Else
RdMove(0).value = True
End If

    Me.Text1.text = IIf(IsNull(rs("foxy_no").value), "", rs("foxy_no").value)
    Me.TXT_order_no.text = IIf(IsNull(rs("order_no").value), "", rs("order_no").value)
    dcBranch.BoundText = IIf(IsNull(rs("branch_no").value), "", rs("branch_no").value)

    TXT_A_NoteID.text = IIf(IsNull(rs("A_NoteID").value), "", val(rs("A_NoteID").value))

    XPTxtID.text = IIf(IsNull(rs("NoteID").value), "", val(rs("NoteID").value))
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    XPTxtVal.text = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
    XPMTxtRemarks.text = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
    txtto.text = IIf(IsNull(rs("too").value), "", rs("too").value)
    txt_general_des.text = IIf(IsNull(rs("general_des").value), "", rs("general_des").value)

    XPDtbTrans.value = IIf(IsNull(rs("NoteDate").value), Date, rs("NoteDate").value)
    XPCboExpensesType.BoundText = IIf(IsNull(rs("ExpensesID").value), "", rs("ExpensesID").value)

    DcFixedAssets.BoundText = IIf(IsNull(rs("FAID").value), "", rs("FAID").value)


    If rs("FAVType").value = 0 Then
        Me.CboType.ListIndex = 0
    ElseIf rs("FAVType").value = 1 Then
        Me.CboType.ListIndex = 1
    End If

    If (rs("bill_Type").value) = 0 Then
        Me.CboPaymentType1.ListIndex = 0
    ElseIf (rs("bill_Type").value) = 1 Then
        Me.CboPaymentType1.ListIndex = 1
    ElseIf (rs("bill_Type").value) = 2 Then
        Me.CboPaymentType1.ListIndex = 2

    End If

    CboPaymentType1_Change
    TXTBankName.text = IIf(IsNull(rs("BankName").value), "", Trim(rs("BankName").value))

    If IsNull(rs("NoteCashingType").value) Then
        Me.CboPayMentType.ListIndex = -1
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        DCVendor.BoundText = ""
        Me.DCChequeBox.BoundText = ""
    
    ElseIf rs("NoteCashingType").value = 0 Then
        Me.CboPayMentType.ListIndex = 0
        Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), 0, rs("BoxID").value)
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        DCVendor.BoundText = ""
    
    ElseIf rs("NoteCashingType").value = 1 Then
        Me.CboPayMentType.ListIndex = 1
        Me.DcboBox.BoundText = ""

        If SystemOptions.ChequeBox = True Then
    
        Else
            Me.DcboBankName.BoundText = rs("BankID").value
        End If
    
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        DCVendor.BoundText = ""
    
        If SystemOptions.ChequeBox = True Then
            Me.DCChequeBox.BoundText = rs("ChequeBoxID").value
        Else
            Me.DCChequeBox.BoundText = ""
            Me.DcboBankName.BoundText = rs("BankID").value
        End If
    
    ElseIf rs("NoteCashingType").value = 2 Then
        Me.CboPayMentType.ListIndex = 2
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
    
        Me.DCVendor.BoundText = rs("CusID").value

    ElseIf rs("NoteCashingType").value = 3 Then
        Me.CboPayMentType.ListIndex = 3
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        DCVendor.BoundText = ""
    ElseIf rs("NoteCashingType").value = 4 Then
        Me.CboPayMentType.ListIndex = 4
        Me.DCAccounts1.BoundText = IIf(IsNull(rs("AccountCode").value), "", rs("AccountCode").value)
        DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = ""
        Me.TxtChequeNumber.text = ""
        DCVendor.BoundText = ""
    
    ElseIf rs("NoteCashingType").value = 5 Then
        Me.CboPayMentType.ListIndex = 5
        Me.DcboBox.BoundText = ""
        Me.DcboBankName.BoundText = rs("BankID").value
        Me.TxtChequeNumber.text = rs("ChqueNum").value
        Me.DtpChequeDueDate.value = rs("DueDate").value
        DCVendor.BoundText = ""
    
    End If

    CboPayMentType_Change

    'ÿMe.DcboBox.BoundText = IIf(IsNull(Rs("BoxID").value), "", Rs("BoxID").value)
    'DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", Val(Me.DcboBox.BoundText))

    If rs("NoteCashingType").value = 0 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    ElseIf rs("NoteCashingType").value = 1 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("BanksData", "BankID", val(Me.DcboBankName.BoundText))
    ElseIf rs("NoteCashingType").value = 2 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DCVendor.BoundText))
    End If

    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.Txt_Numorder.text = IIf(IsNull(rs("NumOrderInpot").value), "", rs("NumOrderInpot").value)
    Me.TxtSerial.text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
    Me.TxtSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
''''''''''''''''''
Me.CboType2.ListIndex = IIf(IsNull(rs("FATypeOp").value), -1, rs("FATypeOp").value)
Me.DcbGroup.BoundText = IIf(IsNull(rs("FAGroupID").value), "", rs("FAGroupID").value)
Me.TxtExcludedValue.text = IIf(IsNull(rs("ExcludedValue").value), 0, rs("ExcludedValue").value)
Me.TxtExcludedValuePrt.text = IIf(IsNull(rs("ExcludedValuePrt").value), 0, rs("ExcludedValuePrt").value)
Me.TxtExcludedValueFixed.text = IIf(IsNull(rs("ExcludedValueFixed").value), 0, rs("ExcludedValueFixed").value)
Me.TxtExcludedValueInst.text = IIf(IsNull(rs("TxtExcludedValueInst").value), 0, rs("TxtExcludedValueInst").value)
Me.TxtNewNuminstallmRemin.text = IIf(IsNull(rs("NewNuminstallmRemin").value), 0, rs("NewNuminstallmRemin").value)
Me.TxtLoseProfitValue.text = IIf(IsNull(rs("LoseProfitValue").value), 0, rs("LoseProfitValue").value)
Me.TxtPurchasePrice.text = IIf(IsNull(rs("PurchasePrice").value), 0, rs("PurchasePrice").value)
Me.TxtAccDepre.text = IIf(IsNull(rs("AccDepre").value), 0, rs("AccDepre").value)
Me.txtCurrentValue.text = IIf(IsNull(rs("CurrentValue").value), 0, rs("CurrentValue").value)
    TxtFASalesPrice.text = IIf(IsNull(rs("FASalesPrice").value), "", rs("FASalesPrice").value)
    
Me.TxtNuminstallmTotal.text = IIf(IsNull(rs("NuminstallmTotal").value), 0, rs("NuminstallmTotal").value)
Me.TxtNuminstallmExcu.text = IIf(IsNull(rs("NuminstallmExcu").value), 0, rs("NuminstallmExcu").value)
Me.TxtNuminstallmRemin.text = IIf(IsNull(rs("NuminstallmRemin").value), 0, rs("NuminstallmRemin").value)
Me.TxtNuminstallmCurr.text = IIf(IsNull(rs("NuminstallmCurr").value), 0, rs("NuminstallmCurr").value)
Me.TxtNewAssest.text = IIf(IsNull(rs("NewAssest").value), "", rs("NewAssest").value)
Me.TxtNewCurrentValue.text = IIf(IsNull(rs("NewCurrentValue").value), 0, rs("NewCurrentValue").value)
Me.txtCountF.text = IIf(IsNull(rs("CountF").value), 0, rs("CountF").value)

If Not (IsNull(rs("ExcludedType").value)) Then
If rs("ExcludedType").value = 1 Then
RdExcluded(1).value = True
Else
RdExcluded(0).value = True
End If
End If

'''''''''''
    Me.dcproject.BoundText = IIf(IsNull(rs("project_Expensen_account").value), "", rs("project_Expensen_account").value)

    If CboPaymentType1.ListIndex = 1 Then 'Õ”«Ì« 

        StrSQL = "SELECT     TOP 100 PERCENT dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, "
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.UserID , dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[value],dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description"
        StrSQL = StrSQL + " FROM         dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN"
        StrSQL = StrSQL + " dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
        StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_id = " & val(rs("A_NoteID").value) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"

        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsDev.RecordCount > 0 Then
            RsDev.MoveFirst
        End If
    
        With Me.VSFlexGrid1
 
            .rows = .FixedRows + RsDev.RecordCount
 
            For i = .FixedRows To .rows
                .TextMatrix(i, .ColIndex("LineNo")) = i
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
            
                .TextMatrix(i, .ColIndex("account_serial")) = IIf(IsNull(RsDev("account_serial").value), "", RsDev("account_serial").value)
            
                If SystemOptions.UserInterface = ArabicInterface Then
            
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_Name").value), "", RsDev("Account_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsDev("Account_NameEng").value), "", RsDev("Account_NameEng").value)
                End If
        
                .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
            
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
            
                RsDev.MoveNext
            Next i
    
        End With

        Exit Sub
    End If

    '-----------------------------------------------------------------------------
    If SystemOptions.SysAppAccoutingType = CompeleteAccounting Then '«·«’Ê·
        '   StrSQL = "Select * From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & Val(Me.XPTxtID.text)
        '   StrSQL = StrSQL + " Order By DEV_ID_Line_No "
        ' StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.*,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.ACCOUNTS.Account_Name FROM    dbo.DOUBLE_ENTREY_VOUCHERS INNER JOIN dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code WHERE     dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID =" & Val(Me.XPTxtID.text) & "Order By DEV_ID_Line_No"

        'StrSQL = "SELECT   dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode,   dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No,dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit , dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID ,dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description  FROM         dbo.ACCOUNTS INNER JOIN dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code"
        'StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit=0  and dbo.DOUBLE_ENTREY_VOUCHERS.notes_all =" & Val(Me.XPTxtID.text) & ") "
        'StrSQL = StrSQL + "ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
        StrSQL = "SELECT  dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetbranch_id , dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetgroupid, dbo.DOUBLE_ENTREY_VOUCHERS.FixedAssetID ,  dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.ACCOUNTS.Account_Name,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value],"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID AS Expr1,"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description , dbo.Notes.order_no"
        StrSQL = StrSQL + " FROM         dbo.ACCOUNTS INNER JOIN"
        StrSQL = StrSQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code INNER JOIN"
        StrSQL = StrSQL + " dbo.Notes ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID"
        StrSQL = StrSQL + " Where (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) And (dbo.DOUBLE_ENTREY_VOUCHERS.notes_all = " & val(Me.XPTxtID.text) & ")"
        StrSQL = StrSQL + " ORDER BY dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No"
    
        Set RsDev = New ADODB.Recordset
        RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsDev.BOF Or rs.EOF) Then
            Me.LblDevID.Caption = RsDev("Double_Entry_Vouchers_ID").value
            Me.lbl(12).Caption = RsDev("Account_Interval_ID").value
            RsDev.MoveFirst

            For i = 1 To RsDev.RecordCount

                If RsDev("Credit_Or_Debit").value = 0 Then
                    Me.DcboDebitSide.BoundText = RsDev("Account_Code").value
                ElseIf RsDev("Credit_Or_Debit").value = 1 Then
                    Me.DcboCreditSide.BoundText = RsDev("Account_Code").value
                End If

                RsDev.MoveNext
            Next i
    
            RsDev.MoveFirst
    
            With Me.VSFlexGrid2

                If Me.dcproject.BoundText = "" Then
                    .rows = .FixedRows + RsDev.RecordCount
                Else
                    .rows = .FixedRows + RsDev.RecordCount - 1
                End If

                For i = .FixedRows To .rows - 1
                    .TextMatrix(i, .ColIndex("LineNo")) = IIf(IsNull(RsDev("DEV_ID_Line_No").value), "", RsDev("DEV_ID_Line_No").value)
            
                    .TextMatrix(i, .ColIndex("LineNo1")) = IIf(IsNull(RsDev("DEV_ID_Line_No1").value), "", RsDev("DEV_ID_Line_No1").value)
            
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("FixedAssetId").value), "", RsDev("FixedAssetId").value)
            
                    .TextMatrix(i, .ColIndex("AccountName")) = getFixedAsstName(val(.TextMatrix(i, .ColIndex("id"))), "name")
           
                    .TextMatrix(i, .ColIndex("groupid")) = IIf(IsNull(RsDev("FixedAssetgroupid").value), "", RsDev("FixedAssetgroupid").value)
            
                    .TextMatrix(i, .ColIndex("branch_id")) = IIf(IsNull(RsDev("FixedAssetbranch_id").value), "", RsDev("FixedAssetbranch_id").value)
                    
                    .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(RsDev("Account_Code").value), "", RsDev("Account_Code").value)
       
                    .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsDev("Double_Entry_Vouchers_Description").value), "", RsDev("Double_Entry_Vouchers_Description").value)
        
                    .TextMatrix(i, .ColIndex("value")) = IIf(IsNull(RsDev("Value").value), "", RsDev("Value").value)
 
                    RsDev.MoveNext
                Next i

                Me.XPTxtVal.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .rows - 1, .ColIndex("Value"))
                ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
                '  Me.TxtTotalCredit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CreditValue"), _
                '  .Rows - 1, .ColIndex("CreditValue"))
                '  Me.TxtTotalDebit.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("DebitValue"), _
                '  .Rows - 1, .ColIndex("DebitValue"))
            End With

        End If

    End If
TxtFASalesPrice_Change
    '-----------------------------------------------------------------------------
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    ReLineGrid
    Me.TxtModFlg = "R"

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
    If CboType.ListIndex = 1 Then
        Dim s As String
        Dim rsDummyAcc As New ADODB.Recordset
        
        s = "Select A211,a24 from branches "
        rsDummyAcc.Open s, Cn, adOpenStatic, adLockReadOnly
        DcboCreditSide.BoundText = rsDummyAcc!A211 & ""
        
        rsDummyAcc.Close
        s = "SELECT group_id FROM FixedAssets Where Id = " & val(DcFixedAssets.BoundText)
         rsDummyAcc.Open s, Cn, adOpenStatic, adLockReadOnly
         
         Dim AccountName As String
    Dim Percentage1 As Integer
    Dim Percentage2 As Integer
    Dim DepType As Integer
    Dim Account_code As String
    Dim Account_code1 As String
    Dim Account_code2 As String
    Dim Account_code3 As String
    Dim Account_code4 As String
    'Â‰« »ÌÃÌ» Õ”«»«  «·„Ã„Ê⁄Â Ê‰”» «·«Â·«ﬂ Ê»ÌÕ”» ⁄„—  «·«’· »«·‘Â—
    GetFixedAssetsGroupAccount val(rsDummyAcc!group_id & ""), , val(Me.dcBranch.BoundText), , , Percentage1, Percentage2, DepType, Account_code, Account_code1, Account_code2, Account_code3, Account_code4
    
        Dim mAccount As String
         mAccount = Account_code
        
    End If
If DcboCreditSide.text = "" And val(TxtFASalesPrice.text) > 0 And CboPayMentType.ListIndex = 0 Then
  If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·Œ“Ì‰… «Ê ‰Ê⁄ «·œ›⁄..!!"
            Else
                Msg = "Select Asset..!!"
            End If
               MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            CboPayMentType.SetFocus
            Sendkeys "{F4}"
            Exit Sub
    End If

        If Trim(Me.DcFixedAssets.BoundText) = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— «·«’·..!!"
            Else
                Msg = "Select Asset..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcFixedAssets.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
                            
        If CheckFixedAssetsDipre(val(DcFixedAssets.BoundText)) = True And Me.TxtModFlg = "N" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "     „ «· Œ·’ „‰ Â–« «·«’· ”«»ﬁ«..!!"
            Else
                Msg = "  Asset was disposed..!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '                                DcFixedAssets.SetFocus
            Sendkeys "{F4}"
            Exit Sub

        End If

        If Me.CboPaymentType1.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·›« Ê—… ...!!!"
            Else
                Msg = "Select Bill Type ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboType.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·”‰œ ...!!!"
            Else
                Msg = "Select   Type ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboType.SetFocus
            Exit Sub
        End If
    
        If Me.CboType.ListIndex = 0 Then '»Ì⁄ «’·
      
            If val(TxtFASalesPrice.text) = 0 Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ»   «œŒ«· ﬁÌ„… «·»Ì⁄ ...!!!"
                Else
                    Msg = "    Enter Price ...!!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                CboType.SetFocus
                Exit Sub
            End If
    
        ElseIf Me.CboType.ListIndex = 1 Then ' Œ—Ìœ «’·
         '  TxtFASalesPrice.Text = 350
        End If
    
        If Me.CboPayMentType.ListIndex = -1 And Me.CboType.ListIndex = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
            Else
                Msg = "Select Payment method ...!!!"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboPayMentType.SetFocus
            Exit Sub
        End If
    
        If Me.CboPayMentType.ListIndex = 2 Then
            If Trim(Me.DCVendor.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
                Else
                    Msg = "Select vendor..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCVendor.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        
        End If
        
        If Me.CboPayMentType.ListIndex = 4 Then
            If Trim(Me.DCAccounts1.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Õ”«»..!!"
                Else
                    Msg = "Select Account..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DCAccounts1.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If
        
        End If
    
        If Me.CboPayMentType.ListIndex = 0 Then
            If Trim(Me.DcboBox.BoundText) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
                Else
                    Msg = "Select Box..!!"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBox.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            '                                                             If Me.DcboBankName.BoundText = "" Then
            '                                                                         If SystemOptions.UserInterface = ArabicInterface Then
            '                                                                             Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
            '                                                                         Else
            '                                                                         Msg = "Select Bank...!!"
            '
            '                                                                        End If
            '                                                                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            ' '                                                                       DcboBankName.SetFocus
            ' '                                                                       SendKeys "{F4}"
            '                                                                        Exit Sub
            '                                                            End If
            '                If Trim$(Me.TxtChequeNumber.text) = "" Then
            '                                      If SystemOptions.UserInterface = ArabicInterface Then
            '                                          Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
            '                                      Else
            '                                      Msg = "Enter Cheque No:...!!"
            '                                      End If
            '                  MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '                  TxtChequeNumber.SetFocus
            '                  Exit Sub
            '              End If
            '
      
            If SystemOptions.ChequeBox = True Then
         
                If DCChequeBox.BoundText = "" Then
                                                           
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "Õœœ Õ«›Ÿ… «·‘Ìﬂ«  ...!!"
                    Else
                        Msg = "Select Cheque Box ...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DCChequeBox.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                                                   
                End If
                        
                If TXTBankName.text = "" Then
                                                       
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "«ﬂ » «”„ »‰ﬂ «·‘Ìﬂ    « ...!!"
                    Else
                        Msg = " Enter Bank Name For Cheque  ...!!"
                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TXTBankName.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                                                
                End If
                            
                If Trim$(Me.TxtChequeNumber.text) = "" Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If

            Else
       
                If Me.DcboBankName.BoundText = "" Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    DcboBankName.SetFocus
                    Sendkeys "{F4}"
                    Exit Sub
                End If

                If Trim$(Me.TxtChequeNumber.text) = "" Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    TxtChequeNumber.SetFocus
                    Exit Sub
                End If
            End If

        ElseIf Me.CboPayMentType.ListIndex = 3 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
                                        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·ÕÊ«·…...!!"
                Else
                    Msg = "Enter Transfer No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
                                                    
            End If
       
        ElseIf Me.CboPayMentType.ListIndex = 5 Then

            If Me.DcboBankName.BoundText = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
                Else
                    Msg = "Select Bank...!!"
                                        
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                DcboBankName.SetFocus
                Sendkeys "{F4}"
                Exit Sub
            End If

            If Trim$(Me.TxtChequeNumber.text) = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
                Else
                    Msg = "Enter Cheque No:...!!"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtChequeNumber.SetFocus
                Exit Sub
                                                    
            End If
       
        End If
    
        If Me.TxtModFlg.text = "N" Then
            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value) = False Then
                        Exit Sub
                    End If
                End If
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If Me.CboPayMentType.ListIndex = 0 Then
                If val(Me.DcboBox.BoundText) <> 0 Then
                    If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtVal.text), XPDtbTrans.value, , , val(Me.XPTxtID.text)) = False Then
                        Exit Sub
                    End If
                End If
            End If
        End If
    
        Dim xrow As Integer

        Dim i As Integer

        ' calcnets

        '-------------------------------------------------------------------------------------------
        Dim notes_result As String
        Dim Vchr_result As String

        '-------------------------------------------------------------------------------------------
        If TxtSerial1.text = "" Then
            Vchr_result = Voucher_coding(val(my_branch), XPDtbTrans.value, 28, 8028)

            If Vchr_result = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   Œ·’ „‰ «’· ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                    MsgBox " Cant't Create  Disposal Of FA  Voucher to this Process no You exceed the maximum number ": Exit Sub
                End If

            Else
         
                If Vchr_result = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    Else
                        MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                    End If

                Else
                    TxtSerial1.text = Vchr_result
                End If
            End If
        End If
    
        If TxtSerial.text = "" Then
            notes_result = Notes_coding(val(my_branch), XPDtbTrans.value)

            If notes_result = "error" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                Else
                    MsgBox " Cant't Create Journal Entry to this Process no You exceed the maximum number ": Exit Sub
                End If

            Else
         
                If notes_result = "" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                    Else
                        MsgBox "You must Define JE Coding ": Exit Sub
                    End If

                Else
                    TxtSerial.text = notes_result
                End If
            End If
        End If
    
        Cn.BeginTrans
        BeginTrans = True
    
        '///////////////NOTESALL
        Dim A_NoteID As Long

        If TxtModFlg.text = "N" Then
            XPTxtID.text = CStr(new_id("notes_all", "NoteID", "", True))
            Me.TxtNoteSerial.text = CStr(new_id("notes_all", "NoteSerial", "", True, "NoteType=80"))
            rs.AddNew
      
        ElseIf Me.TxtModFlg.text = "E" Then
    
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where notes_all=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If
    
        '  Rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
        
    rs("NoteID").value = val(XPTxtID.text)
    rs("YearMove").value = val(TxtYearMove.text)
    rs("YearNotMove").value = val(TxtYearNotMove.text)
    rs("MonthMove").value = val(TxtMonthMove.text)
    rs("MonthNotMove").value = val(TxtMonthNotMove.text)
    rs("AccountCodeVat").value = Me.AccountVat.BoundText
    rs("FATValue").value = val(TxtFATValue.text)
    rs("BillVAT").value = val(TxtBillVAT.text)
    rs("PeriodYear").value = val(TxtPeriodYear.text)
    rs("PeriodMonth").value = val(TxtPeriodMonth.text)
    rs("AgeYear").value = val(TxtAgeYear.text)
    rs("AgeMonth").value = val(TxtAgeMonth.text)
    rs("DpPurchaseDate") = DpPurchaseDate.value
    rs("MonthVAT").value = val(TxtMonthVAT.text)
    rs("NetMonth").value = val(TxtNetMonth.text)
    If RdMove(1).value = True Then
    rs("AsstMove").value = 1
    Else
    rs("AsstMove").value = 0
    End If
        rs("bill_Type").value = Me.CboPaymentType1.ListIndex
        rs("FAVType").value = Me.CboType.ListIndex
        rs("FAID").value = val(DcFixedAssets.BoundText)
        rs("FASalesPrice").value = val(TxtFASalesPrice.text)
    
        rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        rs("foxy_no").value = val(Text1.text)
        rs("order_no").value = TXT_order_no.text
        rs("branch_no").value = val(Me.dcBranch.BoundText)

        ' rs("Note_Value").value = IIf(XPTxtVal.text = "", Null, XPTxtVal.text)
        'rs("Remark").value=""     If SystemOptions.ChequeBox = True Then
        If SystemOptions.ChequeBox = True Then
            rs("ChequeBoxID").value = IIf(DCChequeBox.BoundText = "", Null, DCChequeBox.BoundText)
        Else
            rs("ChequeBoxID").value = Null
                
        End If
                
        If SystemOptions.UserInterface = ArabicInterface Then
            txtmyDes = CboType.text & "   " & DcFixedAssets.text & " » «—ÌŒ " & XPDtbTrans.value
        Else
            txtmyDesE = CboType.text & "    " & DcFixedAssets.text & " Date " & XPDtbTrans.value
        End If
    
        rs("too").value = IIf(txtto.text = "", "", Trim(txtto.text))
        rs("general_des").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
    
        rs("BankName").value = IIf(TXTBankName.text = "", "", Trim(TXTBankName.text))
        rs("CusID").value = Null
        rs("NoteType").value = 8028
        rs("NoteDate").value = XPDtbTrans.value
        rs("UserID").value = user_id
        rs("ExpensesID").value = IIf(XPCboExpensesType.text = "", Null, XPCboExpensesType.BoundText)
  
        Dim bankDes As String
    
        If Me.CboPayMentType.ListIndex = 0 Then
            rs("BoxID").value = val(DcboBox.BoundText)
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 0
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 1

            If SystemOptions.UserInterface = ArabicInterface Then
                bankDes = "  ’—› »‘Ìﬂ —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
            Else
                bankDes = "  Check No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
        
            End If
        
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            rs("NoteCashingType").value = 2
            rs("CusID").value = val(Me.DCVendor.BoundText)
        ElseIf Me.CboPayMentType.ListIndex = 3 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 3

            If SystemOptions.UserInterface = ArabicInterface Then
                bankDes = "  ’—› »ÕÊ«·…  —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
            Else
                bankDes = "  Bank Transfere No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
            End If
    
        ElseIf Me.CboPayMentType.ListIndex = 4 Then
            rs("BoxID").value = Null
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = 4
        
            rs("AccountCode").value = (Me.DCAccounts1.BoundText)
    
        ElseIf Me.CboPayMentType.ListIndex = 5 Then
            rs("BoxID").value = Null
            rs("BankID").value = val(Me.DcboBankName.BoundText)
            rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            rs("DueDate").value = Me.DtpChequeDueDate.value
            rs("NoteCashingType").value = 5

            If SystemOptions.UserInterface = ArabicInterface Then
                bankDes = "  Õ’·  »‘Ìﬂ   —ﬁ„  " & TxtChequeNumber.text & "  ⁄·Ï »‰ﬂ  " & DcboBankName.text
            Else
                bankDes = "  Cheque   No:  " & TxtChequeNumber.text & "  Bank:  " & DcboBankName.text
            End If

            '
        End If
    
        If CboType.ListIndex = 1 Then
            rs("BoxID").value = Null
            rs("BankID").value = Null
            rs("ChqueNum").value = Null
            rs("DueDate").value = Null
            rs("NoteCashingType").value = -1
         
        End If
        rs.Fields("Vat2").value = val(txtVat2.text)
        rs.Fields("Net").value = val(txtNet.text)
    
        rs("project_Expensen_account").value = IIf(Me.dcproject.BoundText = "", "", Me.dcproject.BoundText)
        rs("NumOrderInpot").value = IIf(Trim$(Me.Txt_Numorder.text) = "", Null, Trim$(Me.Txt_Numorder.text))
        rs("Buy").value = "0"
        rs("Remark").value = IIf(txt_general_des.text = "", "", Trim(txt_general_des.text))
        rs("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
        rs("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”·   ›« Ê—…
        rs("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        rs("numbering_type1").value = sand_numbering_type(28) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
     
        rs("sanad_year").value = year(XPDtbTrans.value)
        rs("sanad_month").value = Month(XPDtbTrans.value)

        If dcproject.BoundText <> "" Then
            ' rs("note_value_by_characters").value = Trim$(Me.LblValue.Caption)
        Else
            ' rs("note_value_by_characters").value = WriteNo(Format(Val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0)
        End If

        If Me.TxtModFlg.text = "N" Then
            A_NoteID = CStr(new_id("Notes", "NoteID", "", True))
            TXT_A_NoteID.text = A_NoteID
        Else
            A_NoteID = val(TXT_A_NoteID.text)
        End If
    
        rs("A_NoteID").value = val(A_NoteID)
      ''''''''''''''''''''''''''''''''''''''''''''''''''''
      If RdExcluded(1).value = True Then
      rs("ExcludedType").value = 1
      Else
      rs("ExcludedType").value = 0
      End If
      rs("FATypeOp").value = val(CboType2.ListIndex)
      rs("FAGroupID").value = val(Me.DcbGroup.BoundText)
      rs("ExcludedValue").value = val(TxtExcludedValue.text)
      rs("ExcludedValuePrt").value = val(TxtExcludedValuePrt.text)
      rs("ExcludedValueFixed").value = val(TxtExcludedValueFixed.text)
      rs("TxtExcludedValueInst").value = val(TxtExcludedValueInst.text)
      rs("NewNuminstallmRemin").value = val(TxtNewNuminstallmRemin.text)
      rs("LoseProfitValue").value = val(TxtLoseProfitValue.text)
      rs("PurchasePrice").value = val(TxtPurchasePrice.text)
      rs("AccDepre").value = val(TxtAccDepre.text)
      rs("CurrentValue").value = val(txtCurrentValue.text)
      rs("NuminstallmTotal").value = val(TxtNuminstallmTotal.text)
      rs("NuminstallmExcu").value = val(TxtNuminstallmExcu.text)
      rs("NuminstallmRemin").value = val(TxtNuminstallmRemin.text)
      rs("NuminstallmCurr").value = val(TxtNuminstallmCurr.text)
      rs("NewAssest").value = (TxtNewAssest.text)
      rs("NewCurrentValue").value = val(TxtNewCurrentValue.text)
      
      'rs("CurrentValue").value = val(TxtNewCurrentValue.Text)
      rs("CountF").value = val(txtCountF.text)
      
        rs.update
        
        Savetemp


        Dim ExpensesID As Double
 
        Dim NoteID As String
    
        '  «·«’Ê· „œÌ‰
    
        '//////////////////////////////////////Notes////////////////////////////////////
        Set RsNotes = New ADODB.Recordset
      '  RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        Set RsDev = New ADODB.Recordset
        'RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
         StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                      StrSQL = "SELECT     * from dbo.DOUBLE_ENTREY_VOUCHERS Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
        
        '«·ÿ—› «·„œÌ‰
 
        line_no = 1
    
        ' «·«’Ê· «·ÿ—› «·„œÌÊ‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
                
        RsNotes.AddNew
        NoteID = CStr(new_id("Notes", "NoteID", "", True))
        RsNotes("NoteID").value = CStr(NoteID)
        RsNotes("branch_no").value = val(Me.dcBranch.BoundText)
 
        '    RsNotes("Note_Value").value = IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0)
        RsNotes("Remark").value = Me.txt_general_des
        RsNotes("foxy_no").value = val(Text1.text)

        If Me.CboPayMentType.ListIndex = 0 Then
            RsNotes("BoxID").value = val(DcboBox.BoundText)
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = 0
        ElseIf Me.CboPayMentType.ListIndex = 1 Then
            RsNotes("BoxID").value = Null
                            
            ' RsNotes("BankID").value = Val(Me.DcboBankName.BoundText)
            If SystemOptions.ChequeBox = False Then
        
                rs("BankID").value = val(Me.DcboBankName.BoundText)
            Else
                rs("BankID").value = Null
            End If
                              
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 1
        ElseIf Me.CboPayMentType.ListIndex = 2 Then
            RsNotes("CusID").value = val(DCVendor.BoundText)
 
        ElseIf Me.CboPayMentType.ListIndex = 3 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 3
                      
        ElseIf Me.CboPayMentType.ListIndex = 4 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = Null
            RsNotes("ChqueNum").value = Null
            RsNotes("DueDate").value = Null
            RsNotes("NoteCashingType").value = Null
                       
        ElseIf Me.CboPayMentType.ListIndex = 5 Then
            RsNotes("BoxID").value = Null
            RsNotes("BankID").value = val(Me.DcboBankName.BoundText)
            RsNotes("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
            RsNotes("DueDate").value = Me.DtpChequeDueDate.value
            RsNotes("NoteCashingType").value = 5
                            
        End If
   
        RsNotes("NoteType").value = 8028
        RsNotes("NoteDate").value = XPDtbTrans.value
        RsNotes("UserID").value = user_id
                
        RsNotes("notes_all").value = Me.XPTxtID.text
        RsNotes("NoteSerial").value = Trim$(Me.TxtSerial.text) '„”·”· «·ﬁÌœ
        RsNotes("NoteSerial1").value = Trim$(Me.TxtSerial1.text) '„”·”· «–‰ «·’—›
        RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
        RsNotes("numbering_type1").value = sand_numbering_type(28) '‰Ê⁄  —ﬁÌ„ ›« Ê—… „«·Ì…
        RsNotes("sanad_year").value = year(XPDtbTrans.value)
        RsNotes("sanad_month").value = Month(XPDtbTrans.value)
        RsNotes("note_value_by_characters").value = Trim$(Me.LblValue.Caption) ' WriteNo(Format(IIf(IsNumeric(XPTxtVal.text), XPTxtVal.text, 0), "0.00"), 0, True, ".")
                
        RsNotes.update
                
        XPTxtVal = 0
    
        txtmyDes = txtmyDes & " " & Me.txt_general_des
        txtmyDesE = txtmyDesE & " " & Me.txt_general_des

     ' «·Ï Õ”«» «·ﬁÌ„… «·„÷«›…
        If val(txtCountF) = 0 Then txtCountF = 1
 Dim Percentg As Double
    Dim mm As String
      PercentgValueAddedAccount_Transec XPDtbTrans.value, 12, 1, mm, Percentg
        '«·ÿ—› «·„œÌÊ‰  «·Õ“Ì‰… «Ê «·»‰ﬂ
           ' If CboType.ListIndex = 1 Then
                If val(TxtExcludedValuePrt.text) > 0 Then
                    RsDev.AddNew
                    RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                    RsDev("DEV_ID_Line_No").value = line_no
                    RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                    RsDev("Account_Code").value = DcboCreditSide.BoundText
                   ' RsDev("Vatyo").value = Percentg
                   ' RsDev("Vat").value = val(txtVat2)
                    
                    RsDev("Value").value = val(TxtExcludedValuePrt)
                    RsDev("Credit_Or_Debit").value = 0
                    RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
                    RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
                    RsDev("RecordDate").value = Me.XPDtbTrans.value
                    RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
                    RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                    RsDev("UserID").value = Me.DCboUserName.BoundText
                    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    RsDev("notes_all").value = Me.XPTxtID.text
                               
                    XPTxtVal = val(XPTxtVal.text) + val(TxtFASalesPrice.text)
                    RsDev.update
                    line_no = line_no + 1
                    
                    
                                        RsDev.AddNew
                    RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
                    RsDev("DEV_ID_Line_No").value = line_no
                    RsDev("DEV_ID_Line_No1").value = setfoxy_Line
                    RsDev("Account_Code").value = mAccount
                   ' RsDev("Vatyo").value = Percentg
                   ' RsDev("Vat").value = val(txtVat2)
                    
                    RsDev("Value").value = val(TxtExcludedValuePrt)
                    RsDev("Credit_Or_Debit").value = 1
                    RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
                    RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
                    RsDev("RecordDate").value = Me.XPDtbTrans.value
                    RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
                    RsDev("branch_id").value = val(Me.dcBranch.BoundText)
                    RsDev("UserID").value = Me.DCboUserName.BoundText
                    RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
                    RsDev("notes_all").value = Me.XPTxtID.text
                               
                    XPTxtVal = val(XPTxtVal.text) + val(TxtFASalesPrice.text)
                    RsDev.update
                    line_no = line_no + 1
                    
                    s = " UPDATE FixedAssets SET CurrentValue =  " & TxtNewCurrentValue & "  WHERE id = " & DcFixedAssets.BoundText
                    Cn.Execute s
                End If

         End If
        
        If val(TxtFASalesPrice.text) > 0 Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = DcboCreditSide.BoundText
            RsDev("Vatyo").value = Percentg
            RsDev("Vat").value = val(txtVat2)
            
            RsDev("Value").value = IIf(IsNumeric(TxtFASalesPrice.text), val(TxtFASalesPrice.text) + val(TxtFATValue.text), 0) + val(txtVat2) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 0
            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
                       
            XPTxtVal = val(XPTxtVal.text) + val(TxtFASalesPrice.text)
            RsDev.update
            line_no = line_no + 1
        End If
     
        If val(TxtAccDepre.text) > 0 And CboType.ListIndex <> 1 Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = get_FixedAsset_Account(group_id, branch_id, "Account_Code2")
            
            RsDev("Value").value = IIf(IsNumeric(TxtAccDepre.text), TxtAccDepre.text, 0) / val(txtCountF)  '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 0
            RsDev("VATYou").value = val(txtVat2)
            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes & CHR(13) & "    «·«Â·«ﬂ "   ' .TextMatrix(I, .ColIndex("des"))
            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
                        
            RsDev.update
            XPTxtVal = val(XPTxtVal.text) + val(TxtAccDepre.text)
            line_no = line_no + 1
        End If
          
        If val(TxtFASalesPrice) > val(TxtNewCurrentValue.text) Then
            ProfitOrLose = 1
            ProfitOrLoseValue = val(TxtFASalesPrice) - val(TxtNewCurrentValue.text)
        ElseIf val(TxtFASalesPrice) < val(TxtNewCurrentValue.text) Then
            ProfitOrLose = 0
            ProfitOrLoseValue = Abs(val(TxtNewCurrentValue.text) - val(TxtFASalesPrice))
            XPTxtVal = val(XPTxtVal.text) + (val(TxtNewCurrentValue.text) - val(TxtFASalesPrice))
        Else
            ProfitOrLose = -1
            ProfitOrLoseValue = 0
        End If
          
        If val(ProfitOrLoseValue) > 0 And CboType.ListIndex <> 1 Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = DCAccounts.BoundText
            RsDev("Value").value = IIf(IsNumeric(ProfitOrLoseValue), ProfitOrLoseValue, 0) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = ProfitOrLose
            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes ' .TextMatrix(I, .ColIndex("des"))
            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
                        
            RsDev.update
          
            line_no = line_no + 1
        End If
          
        If val(TxtPurchasePrice.text) > 0 And CboType.ListIndex <> 1 Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = get_FixedAsset_Account(group_id, branch_id, "Account_Code")
            RsDev("Value").value = IIf(IsNumeric(TxtPurchasePrice.text), TxtPurchasePrice.text, 0) / val(txtCountF) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes & CHR(13) & "   ﬁÌ„… ‘—«¡ «·«’·"   ' .TextMatrix(I, .ColIndex("des"))
            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
       RsDev.update
            line_no = line_no + 1
        End If
        
        
                If val(txtAdditions.text) > 0 And CboType.ListIndex <> 1 Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = get_FixedAsset_Account(group_id, branch_id, "Account_Code")
            RsDev("Value").value = IIf(IsNumeric(txtAdditions.text), txtAdditions.text, 0) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes & CHR(13) & "«÷«›«  ⁄·Ì «·«’·"   ' .TextMatrix(I, .ColIndex("des")) & "«÷«›«  ⁄·Ì «·«’·"
            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
       RsDev.update
            line_no = line_no + 1
        End If
        
        
        If val(TxtFATValue.text) > 0 And AccountVat.BoundText <> "" Then
            RsDev.AddNew
            RsDev("Double_Entry_Vouchers_ID").value = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True) 'LngDevID
            RsDev("DEV_ID_Line_No").value = line_no
            RsDev("DEV_ID_Line_No1").value = setfoxy_Line
            RsDev("Account_Code").value = AccountVat.BoundText
            RsDev("Value").value = IIf(IsNumeric(TxtFATValue.text), val(TxtFATValue.text), 0) '.TextMatrix(I, .ColIndex("VALUE"))
            RsDev("Credit_Or_Debit").value = 1
            RsDev("Double_Entry_Vouchers_Description").value = txtmyDes + "Õ”«» «·ﬁÌ„… «·„÷«›… ·· Œ·’ „‰ «·«’Ê·" ' .TextMatrix(I, .ColIndex("des"))
            RsDev("Double_Entry_Vouchers_Descriptione").value = txtmyDesE + "Account of VAT Dis.Assests"
            RsDev("RecordDate").value = Me.XPDtbTrans.value
            RsDev("Notes_ID").value = val(NoteID) '(XPTxtID.text)
            RsDev("branch_id").value = val(Me.dcBranch.BoundText)
            RsDev("UserID").value = Me.DCboUserName.BoundText
            RsDev("Account_Interval_ID").value = SystemOptions.SysCurrentAccountIntervalID
            RsDev("notes_all").value = Me.XPTxtID.text
                        
            RsDev.update
            line_no = line_no + 1
        End If

  '  End If
        If val(txtVat2) <> 0 Then
            Dim StrAccountCodeCridet As String
            GetValueAddedAccount XPDtbTrans.value, , StrAccountCodeCridet, 1, 12
                line_no = line_no + 1
           LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)  'LngDevID
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(txtVat2), 1, Msg & "    Õ”«»  «·ﬁÌ„… «·„÷«›… ", val(NoteID), , , , XPDtbTrans.value, val(DCboUserName.BoundText), , , , , , 1, , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            line_no = line_no + 1
        End If
    '&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
    '   UpdateFixedAssetPurchaseInformations ' ÕœÌÀ »Ì«‰«  «·«’· «
   ' End If
    LblDevID.Caption = LngDevID
    lbl(12).Caption = SystemOptions.SysCurrentAccountIntervalID
    
ll:

    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    CuurentLogdata

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
        
            End If

            Fg_Journal.Enabled = False

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        
            End If

            Fg_Journal.Enabled = False
    End Select

    'Õ›Ÿ »Ì«‰«  «·‘Ìﬂ« 
    saveChequeBoxContents (val(Me.XPTxtID.text))
      
    TxtModFlg.text = "R"
    Dim sql As String
    sql = "Update notes    set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text)
    Cn.Execute sql
    sql = "Update   notes_all  set note_value_by_characters='" & WriteNo(Format(val(Me.XPTxtVal.text), "0.00"), 0, True, ".", , 0) & "' where NoteSerial=" & val(TxtSerial.text)
    Cn.Execute sql
    If CboType.ListIndex = 1 Then
    
    
        sql = "Select * from TblFixedAssestTmpValue Where TransId = " & val(TxtSerial1)
        Dim rsDummyFi As New ADODB.Recordset
        rsDummyFi.Open sql, Cn, adOpenKeyset, adLockOptimistic
        If rsDummyFi.EOF Then
            rsDummyFi.AddNew
            rsDummyFi!TransID = val(TxtSerial1)
        Else
        End If
        
        rsDummyFi!FixedID = val(DcFixedAssets.BoundText)
        rsDummyFi!netvalue = val(TxtExcludedValueFixed)
        rsDummyFi.update
        rsDummyFi.Close
        
        sql = "Select Sum(netvalue) netvalue from TblFixedAssestTmpValue Where TransId = " & val(TxtSerial1)
        sql = sql & " and FixedID = " & val(DcFixedAssets.BoundText)
        
        rsDummyFi.Open sql, Cn, adOpenKeyset, adLockOptimistic
        Dim mNoteValue As Double
        If Not rsDummyFi.EOF Then
            mNoteValue = val(rsDummyFi!netvalue & "")
            
        Else
            mNoteValue = val(TxtExcludedValueFixed)
        End If
        
        sql = "Update   FixedAssets  set CurrentValue=" & mNoteValue & " where id=" & val(DcFixedAssets.BoundText)
        Cn.Execute sql
    Else
        sql = "Update   FixedAssets  set Status_id='" & CboType.ListIndex + 2 & "' where id=" & val(DcFixedAssets.BoundText)
        Cn.Execute sql
        
            sql = "  update FixedAssets  set   KhordaPrice =0 ,  saleprice=" & val(TxtFASalesPrice.text) & " where id=" & val(DcFixedAssets.BoundText)
    Cn.Execute sql
    End If

updateNotesValueAndNobytext (val(NoteID))
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "cant save " & CHR(13)
            Msg = Msg + "Invalid entry value " & CHR(13)
            Msg = Msg + "Check data and try again"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorr.... Error during saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Sub Savetemp()
    
    
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
 
 
SaveQRCode "notes_all", "NoteID", val(XPTxtID.text), TxtSerial1.text, (XPDtbTrans.value), _
        (txtNet.text), Picture2, 0, (txtVat2.text), (txtNet.text)


End Sub
Function UpdateFixedAssetPurchaseInformations(Optional delete As Boolean)
    Dim sql As String
    Dim i As Integer
    Dim KhordaPrice As Double
    Dim currentvalue As Double
    Dim PurcahsePrice As Double
    Dim Installmentvalue As Double

    With Me.VSFlexGrid2

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
 
                sql = "update FixedAssets set PurchaseDate=CONVERT(DATETIME, '" & XPDtbTrans.value & " 00:00:00', 103), PurchaseBillId=" & TxtSerial1.text & ",PurchasePrice="
           
                PurcahsePrice = val(.TextMatrix(i, .ColIndex("value")))
                sql = sql & PurcahsePrice
           
                Dim noofinstllments As Double
              
                GetAllDataAboutFixedAsset val(.TextMatrix(i, .ColIndex("id"))), , , , , , , , , , , , , noofinstllments, , , , , , KhordaPrice
                currentvalue = PurcahsePrice - KhordaPrice
                sql = sql & ",CurrentValue= " & currentvalue

                If noofinstllments = 0 Then
                    noofinstllments = 0
                Else
                    Installmentvalue = Round(currentvalue / noofinstllments, 2)
                End If
            
                sql = sql & ",Installmentvalue= " & Installmentvalue
                sql = sql & ",NoteSerial=' " & Me.TxtNoteSerial.text & "'"
                sql = sql & "  where id=" & val(.TextMatrix(i, .ColIndex("id")))
                Cn.Execute sql

                If noofinstllments <> 0 Then
                    updateFixedAsseTInstallmentInformations val(.TextMatrix(i, .ColIndex("id"))), , , , XPDtbTrans.value, , , , True, True ' ÕœÌÀ »Ì«‰«  «·«ﬁ”«ÿ
                End If

                If delete = True Then
                    '  sql = "update FixedAssets NoteSerial=0,  PurchaseBillId=" & "" & ",PurchasePrice=0,Installmentvalue=0,CurrentValue=0"
                End If
            
            End If
        
        Next i

    End With

End Function

Public Function save_General_cost_center(cost_center_id As String, _
                                         cost_center, _
                                         opr_type As String, _
                                         record_date As Date) 'As String, value As Double, depit_or_credit As Boolean, opr_id As Double, opr_type As String, account_no As String, account_name As String, line_no As Double, recorddate As Date)
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
  '  rs.Open "marakes_taklefa_temp", Cn, adOpenStatic, adLockOptimistic, adCmdTable
  StrSQL = "SELECT   *  from dbo.marakes_taklefa_temp Where (1 = -1)"
   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
    With Fg_Journal
 
        .rows = .rows + 1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
        
                rs.AddNew
                rs("cost_center_id").value = cost_center_id
                rs("cost_center").value = cost_center
                rs("value").value = .TextMatrix(i, .ColIndex("value"))
                rs("depit_or_credit").value = "„œÌ‰"
                rs("opr_id").value = Me.Text1.text
                rs("kedno").value = Me.Text1.text
                rs("opr_type").value = opr_type
                rs("account_name").value = .TextMatrix(i, .ColIndex("AccountName"))
                rs("account_no").value = .TextMatrix(i, .ColIndex("AccountCode"))
                rs("line_no").value = .TextMatrix(i, .ColIndex("LineNo1"))
                rs("record_date").value = record_date
                rs.update
        
            End If

        Next i

    End With

    rs.Close
End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "NoteID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    On Error GoTo ErrTrap

    If SystemOptions.banks_Accounts3 = True Then
                If ChequeBoxOperations1(val(Me.XPTxtID)) = False Then
                    '         Msg = " ·« Ì„ﬂ‰ «·”„«Õ »Õ–› Â–… «·⁄„·Ì…"
                    '         Msg = Msg & CHR(13) & " ÌÊÃœ ⁄„·Ì… ”œ«œ ··‘Ìﬂ „”Ã·Â "
                    '         MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    '         Exit Sub
                End If
    End If
    
    Dim noOfInstallments As Integer 'Â–« «·Ã“¡ Ì √ﬂœ „‰  ‰›Ì– «ﬁ”«ÿ «Â·«ﬂ
    Dim msgstr As String
    Dim i As Integer

    '    UpdateFixedAssetPurchaseInformations True
    
    If XPTxtID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title) = vbYes Then
            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
 
'            StrSQL = "Delete From notes Where NoteID=" & val(TXT_A_NoteID.text)
            StrSQL = "Delete From notes Where notes_all=" & val(XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From ExpensesDetails Where NoteSerial1='" & val(TxtSerial1.text) & "'"
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            StrSQL = "Delete  TblChecqueBoxContent1  where NoteID =" & val(Me.XPTxtID.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
    
            'UPDATEStatusToNewAsset
            
            
            Dim sql As String
       Dim rsDummyFi As New ADODB.Recordset
            
                If CboType.ListIndex = 1 Then
    
    
      
        
        sql = "Select Sum(netvalue) netvalue from TblFixedAssestTmpValue Where TransId = " & val(TxtSerial1)
        sql = sql & " and FixedID = " & val(DcFixedAssets.BoundText)
        
        rsDummyFi.Open sql, Cn, adOpenKeyset, adLockOptimistic
        Dim mNoteValue As Double
        If Not rsDummyFi.EOF Then
            mNoteValue = val(rsDummyFi!netvalue & "")
            
        Else
            mNoteValue = val(TxtExcludedValueFixed)
        End If
        
            sql = "Update   FixedAssets  set CurrentValue=" & mNoteValue & " where id=" & val(DcFixedAssets.BoundText)
            Cn.Execute sql
            
            sql = "Delete TblFixedAssestTmpValue Where TransId = " & val(TxtSerial1)
            Cn.Execute sql
        End If
 
            sql = "Update   FixedAssets  set Status_id=0 " & " where id=" & val(DcFixedAssets.BoundText)
            Cn.Execute sql
   
            sql = "  update FixedAssets  set KhordaPrice=1,   saleprice=0  where id=" & val(DcFixedAssets.BoundText)
            Cn.Execute sql
  
  
  
  
  
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    Fg_Journal.Clear flexClearScrollable, flexClearEverything
                    Fg_Journal.rows = 3
                    Fg_Journal.Enabled = False
                
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
                    VSFlexGrid1.rows = 2
                    VSFlexGrid1.Enabled = False
                
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate
End Sub

Function FillGridWithData()

End Function

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            
            End If

        Next i

    End With

    IntCounter = 0

    With Me.VSFlexGrid1

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                    
            End If

        Next i

    End With

    IntCounter = 0

    With Me.VSFlexGrid2

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("des")) = " ﬁÌ„… ‘—«¡ «·«’· " & .TextMatrix(i, .ColIndex("AccountName"))
                    
                Else
                    .TextMatrix(i, .ColIndex("des")) = "PURCHASE Value Of Asset " & .TextMatrix(i, .ColIndex("AccountName"))
                End If
                    
            End If

        Next i

    End With

End Sub

Function UPDATEStatusToNewAsset()
    Dim StrSQL As String
    Dim i As Integer
 
    With Me.VSFlexGrid2

        For i = .FixedRows To .rows - 1

            If .TextMatrix(i, .ColIndex("id")) <> "" Then
                StrSQL = "UPDATE FixedAssets SET CurrentValue = 0,PurchaseBillId='',Installmentvalue = 0,NoteSerial='', New_or_opening=0 ,PurchasePrice=0 where  id=" & val(.TextMatrix(i, .ColIndex("id")))
   
                Cn.Execute StrSQL
            End If

        Next i

    End With

End Function

Private Sub PutData()

    'MsgBox Fg_Journal.Row & "---" & Fg_Journal.ColKey(Fg_Journal.Col)
    With Fg_Journal

        If Len(TxtDes.text) > 0 Then
            .cell(flexcpData, .Row, .ColIndex("Des")) = TxtDes.text
            .cell(flexcpPicture, .Row, .ColIndex("Des")) = ImgNote.Picture
            .cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        Else
            .cell(flexcpData, .Row, .ColIndex("Des")) = ""
            .cell(flexcpPicture, .Row, .ColIndex("Des")) = Empty
            .cell(flexcpPictureAlignment, .Row, .ColIndex("Des")) = flexAlignLeftCenter
        End If

    End With

End Sub

Function sand_numbering() As String
    On Error Resume Next
    Dim start_at As Integer
    Dim end_at As Integer
    Dim auto_sanad_no As String
    Dim NO As String
    auto_sanad_no = ""
    departement_name = 1
    branch_no = 1
    connection_string = Cn.ConnectionString
    numbering.ConnectionString = connection_string
    numbering.CommandType = adCmdText
    numbering.RecordSource = "select * from sanad_numbering where branch_no=" & my_branch & " and departement='" & departement_name & "' and  sanad_no=1"
    numbering.Refresh

    If numbering.Recordset.RecordCount = 0 Then
        numbering_type = 0
    Else
        numbering_type = numbering.Recordset.Fields!numbering_id
        start_at = numbering.Recordset.Fields!start_at
        end_at = numbering.Recordset.Fields!end_at

    End If

    If numbering_type = 1 Then
        detect_no.ConnectionString = connection_string
        detect_no.CommandType = adCmdText
        detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type  ' branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type
        detect_no.Refresh

        If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
 
            If end_at = 0 Then end_at = detect_no.Recordset.Fields!last_sand_no + 1
 
            If detect_no.Recordset.Fields!last_sand_no >= end_at Then
                sand_numbering = "error"
                Exit Function
            End If
        End If

    Else

        If numbering_type = 2 Then
 
            detect_no.ConnectionString = connection_string
            detect_no.CommandType = adCmdText
            detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & mId(Format$(Now, "dd/mm/yyyy"), 4, 2)
            'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "' and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "and sanad_month=" & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2)
            detect_no.Refresh

            If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)

                If end_at = 0 Then end_at = NO + 1
                If NO >= end_at Then
                    sand_numbering = "error"
                    Exit Function
                End If
            End If

        Else

            If numbering_type = 3 Then
 
                detect_no.ConnectionString = connection_string
                detect_no.CommandType = adCmdText
                detect_no.RecordSource = "select max(NoteSerial) as last_sand_no from  notes_all where NoteType=80 and numbering_type=" & numbering_type & "and sanad_year=" & mId(Format$(Now, "dd/mm/yyyy"), 7, 4)
                'detect_no.RecordSource = "select max(sanad_no) as last_sand_no from  sandat_ked where  branch_no=" & branch_no & " and departement='" & departement_name & "'  and  type='" & "”‰œ ﬁÌœ" & "' and numbering_type=" & numbering_type & "and sanad_year=" & Mid(Format$(Now, "dd/mm/yyyy"), 7, 4)
                detect_no.Refresh

                If Not IsNull(detect_no.Recordset.Fields!last_sand_no) Then
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)

                    If end_at = 0 Then end_at = NO + 1
                    If NO >= end_at Then
                        sand_numbering = "error"
                        Exit Function
                    End If
                End If
 
            End If
 
        End If
    End If

    If detect_no.Recordset.RecordCount = 0 Or IsNull(detect_no.Recordset.Fields!last_sand_no) Then

        If numbering_type = 0 Then
            ' auto_sanad_no = 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = start_at
            Else
                
                If numbering_type = 2 Then
                    auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & mId(Format$(Now, "dd/mm/yyyy"), 4, 2) & start_at

                Else

                    If numbering_type = 3 Then
                        auto_sanad_no = mId(Format$(Now, "dd/mm/yyyy"), 7, 4) & start_at

                    End If
                End If
            End If
        End If

    Else

        If numbering_type = 0 Then
            'auto_sanad_no = x + 1
        Else

            If numbering_type = 1 Then
                auto_sanad_no = detect_no.Recordset.Fields!last_sand_no + 1
            Else
                
                If numbering_type = 2 Then
                    '  If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 6) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) Then
                    ' no = 1
                    '  auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & Mid(Format$(Now, "dd/mm/yyyy"), 4, 2) & "1"
                    '  Else
                    NO = mId(detect_no.Recordset.Fields!last_sand_no, 7, Len(detect_no.Recordset.Fields!last_sand_no) - 6)
                    auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 6) & (NO + 1)
                    '  End If
                      
                Else

                    If numbering_type = 3 Then
                        '    If Mid(detect_no.Recordset.Fields!last_sand_no, 1, 4) <> Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) Then
                        'no = 1
                        '    auto_sanad_no = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & "1"
                        '    Else
                        NO = mId(detect_no.Recordset.Fields!last_sand_no, 5, Len(detect_no.Recordset.Fields!last_sand_no) - 4)
                        auto_sanad_no = mId(detect_no.Recordset.Fields!last_sand_no, 1, 4) & (NO + 1)

                        '    End If

                    End If
                End If
            End If
        End If

    End If

    sand_numbering = auto_sanad_no

    'MsgBox auto_sanad_no

End Function

Function setfoxy_Line() As Double
    
    Dim X As Double
    X = CStr(new_id("foxy", "id1", "", True))
    setfoxy_Line = X
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "foxy", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    rs("id1").value = X ' last_line_id
 
    rs.update
    
End Function

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "—ﬁ„ «·”‰œ  " & TxtSerial1.text & CHR(13) & "   «· «—ÌŒ  " & XPDtbTrans & CHR(13) & "   «·›—⁄ " & dcBranch & CHR(13) & "   ‰Ê⁄ «·”‰œ " & CboType & CHR(13) & "     «·«’·  " & DcFixedAssets & CHR(13) & "   ÿ—Ìﬁ… «·»Ì⁄  " & CboPayMentType & CHR(13) & "   ﬁÌ„…  «·‘—«¡  " & TxtPurchasePrice & CHR(13) & "„Ã„⁄ «·«Â·«ﬂ " & TxtAccDepre & CHR(13) & "      «·ﬁÌ„… «·œ› —Ì…  " & txtCurrentValue & CHR(13) & "   ﬁÌ„…  «·»Ì⁄  " & TxtFASalesPrice & CHR(13) & "     «·—»Õ «Ê «·Œ”«—…  " & TxtLoseProfitValue & CHR(13) & "   «·Œ“Ì‰… " & DcboBox & CHR(13) & "   «·»‰ﬂ  " & DcboBankName & CHR(13) & "   —ﬁ„ «·‘Ìﬂ " & TxtChequeNumber & CHR(13) & "    «—ÌŒ «·«” Õﬁ«ﬁ  " & DtpChequeDueDate & CHR(13) & "   «·⁄„Ì·  " & DCVendor & CHR(13) & " «·Õ”«»  " & DCAccounts1 & CHR(13) & "   «·‘—Õ «·⁄«„  " & txt_general_des & CHR(13) & "   «Ã„«·Ì «·”‰œ    " & XPTxtValView
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill . No " & TxtSerial1.text & CHR(13) & "   Date  " & XPDtbTrans & CHR(13) & "   Branch " & dcBranch & CHR(13) & "    Type   " & CboType & CHR(13) & "     F.A. Name  " & DcFixedAssets & CHR(13) & "  Salle Type  " & CboPayMentType & CHR(13) & "Purchase Price " & TxtPurchasePrice & CHR(13) & "Acc Depre " & TxtAccDepre & CHR(13) & "Current Value " & txtCurrentValue & CHR(13) & "  Sales Price " & TxtFASalesPrice & CHR(13) & "Lose /Profit Value " & TxtLoseProfitValue & CHR(13) & "   Box " & DcboBox & CHR(13) & "   Bank  " & DcboBankName & CHR(13) & "   Cheque No:   " & TxtChequeNumber & CHR(13) & "   Supplier  " & DCVendor & CHR(13) & " Account  " & DCAccounts1 & CHR(13) & "  Remarks  " & txt_general_des & CHR(13) & "   Vchr Total   " & XPTxtValView
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 8028, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , val(TxtSerial), val(TxtSerial1)
    Else
        AddToLogFile CInt(user_id), 8028, Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TxtSerial, TxtSerial1
    End If
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
        End If
    End If

    If KeyCode = vbKeyF12 Then
        If Cmd(0).Enabled = False Then Exit Sub
        Cmd_Click (0)
    End If

    If KeyCode = vbKeyF11 Then
        If Cmd(1).Enabled = False Then Exit Sub
        Cmd_Click (1)
    End If

    If KeyCode = vbKeyF10 Then
        If Cmd(2).Enabled = False Then Exit Sub
        Cmd_Click (2)
    End If

    If KeyCode = vbKeyF9 Then
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If BolRtl = True Then

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "«·„’—Ê›« ", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    Else

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "Add New Record..." & Wrap & "Shortcut Key F12 OR Enter" & Wrap & "OR Alt+N", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit the Current Record..." & Wrap & "Shortcut Key F11 " & Wrap & "OR Alt+E", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save the New Record OR Save the Editing in the Current Record..." & Wrap & "Shortcut Key F10 " & Wrap & "OR Alt+S", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Cancel the New Record OR Cancel Editing in the Current Record..." & Wrap & "Shortcut Key F9 " & Wrap & "OR Alt+U", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete the Current Record..." & Wrap & "Shortcut Key F8 " & Wrap & "OR Alt+D", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Close this Screen" & Wrap & "OR Alt+X", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "Expenses", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help..." & Wrap & "Display Help for this Screen" & Wrap & "Shortcut Key F1" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                
                End If

        End Select

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPCboExpensesType_Change()

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        Me.DcboDebitSide.BoundText = ModAccounts.GetMyAccountCode("ExpensesType", "ID", val(Me.XPCboExpensesType.BoundText))
    End If

End Sub

Private Sub XPDtbTrans_Change()
    TxtSerial.text = ""
    TxtSerial1.text = ""
    GetDataAssestVAT
End Sub

Private Sub XPTxtVal_Change()
    XPTxtValView.text = Format(val(XPTxtVal.text), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 0)

    Else

        'Me.LblValue.Caption = WriteNo(XPTxtVal.text, 0, , , , 1)
        Me.LblValue.Caption = WriteNo(Format(Me.XPTxtVal.text, "0.00"), 0, True, ".", , 1)

    End If
    
End Sub

Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
    'KeyAscii = KeyAscii_Num(KeyAscii, XPTxtVal.text, 0)
End Sub

Private Sub XPTxtVal_Validate(Cancel As Boolean)
    'If Val(XPTxtVal.Text) = 0 Then
    '    Set TTD = New clstooltipdemand
    '    TTD.Style = TTBalloon
    '    TTD.Icon = TTIconWarning
    '    TTD.Centered = True
    '    TTD.RightToLeft = True
    '    TTD.VisibleTime = 600
    '    TTD.BackColor = 0
    '    TTD.Title = "ﬁÌ„… «·„’—Ê›« "
    '    TTD.TipText = "»—Ã«¡ ﬂ «»… ﬁÌ„… «·„’—Ê›« "
    '    TTD.PopupOnDemand = True
    '    TTD.CreateToolTip XPTxtVal.hwnd
    '    TTD.Show 0, XPTxtVal.Height / Screen.TwipsPerPixelX - 1    '//In Pixel only
    '    Cancel = True
    'Else
    '    TTD.Destroy
    'End If
End Sub

Private Sub ViewDataList()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.VSFlexGrid
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    'Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.VSFlexGrid

    With FG
        .Cols = 18
        .RowHeightMin = 320
        .ExplorerBar = flexExSortShowAndMove
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        StrSQL = "SELECT NoteID, NoteSerial, NoteDate, Name, Note_Value, BoxName," & "Remark, UserName From ExpensesReport"
        StrSQL = StrSQL + " Order By NoteID"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
        'Â‰« Ìﬂ » ﬂÊœ ·⁄„· „⁄œ·  Õ„Ì· «·»Ì«‰« 
        '------------------------------------
        '
        '
        '
        '
    
        '------------------------------------
        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·⁄„·Ì…"
        .ColKey(0) = "NoteID"
        .TextMatrix(0, 1) = "ﬂÊœ «·⁄„·Ì…"
        .ColKey(1) = "NoteSerial"
        .TextMatrix(0, 2) = "«· «—ÌŒ"
        .ColKey(2) = "NoteDate"
        .TextMatrix(0, 3) = "‰Ê⁄ «·„’—Ê›« "
        .ColKey(3) = "Name"
        .TextMatrix(0, 4) = "ﬁÌ„… «·„’—Ê›« "
        .ColKey(4) = "Note_Value"
        .ColFormat(.ColIndex("Note_Value")) = "#,###.##"
        .TextMatrix(0, 5) = "«”„ «·Œ“‰…"
        .ColKey(5) = "BoxName"
        .TextMatrix(0, 6) = "„·«ÕŸ« "
        .ColKey(6) = "Remark"
        .TextMatrix(0, 7) = "Õ—— »Ê«”ÿ…"
        .ColKey(7) = "UserName"
    
        'Rs.Close
        'Set Rs = Nothing
        .AutoSize 0, .Cols - 1, False
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.VSFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "Note_Value"
    FrmView.vsfGroup1.sql = StrSQL
    FrmView.vsfGroup1.ShowTreeGroups = True
    FrmView.vsfGroup1.update
    FrmView.SetDblClickRetrun Me, "NoteID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·»Ì«‰«  «·„’—Ê›« "
    FrmView.show
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    'LblValue.Visible = False
 CmdAttach.Caption = "Attachments"
 Me.lbl(42).Caption = "Purchase Date"
    lbl(24).Caption = "Hint."
    lbl(25).Caption = "This Window Allow Disposal Of Fixed Assets"
lbl(34).Caption = "Opr Type"
Frame2.Caption = "Current Data"
Frame5.Caption = "Disposal Date"
    
    lbl(37).Caption = " Total Installments"
lbl(36).Caption = " Remains Installments"
Label12.Caption = "Remain Value"
lbl(35).Caption = " Exe Installments"
lbl(38).Caption = "Installment Value"
 
 
Frame6.Caption = "Partial disposal Data"
Label14.Caption = "New Name"
Label15.Caption = "Group"
lbl(39).Caption = "Current Value"
lbl(40).Caption = "Remain Inst."
 lbl(43).Caption = "Cheque Box"
 lbl(33).Caption = "Account"
ISButton1.Caption = "Attachments"


Label9.Caption = "Disposal Value"
Label11.Caption = "Disposal part Value"
Label13.Caption = "New InstallmentsValue"


Frame4.Caption = "Sale Price"

    lbl(23).Caption = " Type"
    Label3.Caption = "GL No."
    lbl(14).Caption = "Project#"
    'Label1.Caption = "Manual #"
    Me.ALLButton1.Caption = "Cost Center"
    lbl(15).Caption = "Sales Method"
    lbl(16).Caption = "Box Name"
    lbl(20).Caption = "General Des"
    lbl(21).Caption = "Order No:"
    Label1.Caption = "Branch"
    lbl(26).Caption = "Account"
    lbl(28).Caption = "Purch. Price"
    lbl(29).Caption = "Acc Dep"
    lbl(30).Caption = "Current Value"
    lbl(31).Caption = "Sales Value"
    lbl(32).Caption = "Profit Or Loss"

    lbl(26).Caption = "ACC."

    Label8.Caption = "General C. C."

    With Me.CboPayMentType
        .Clear
        .AddItem "Cash"
        .AddItem "Cheque"
        .AddItem "Credit"
        .AddItem "Transfer"
        .AddItem "Account"
        .AddItem "Collected Cheque"
    End With

    With Me.CboPaymentType1
        .Clear
        .AddItem "Expenses"
        .AddItem "Accounts"
        .AddItem "Fixed Asset Purchase"
    End With

    With Me.CboType
        .Clear
        .AddItem "Disposal Of FA"
        .AddItem "Exclusion Of FA"
     
    End With

    With Me.CboType2
        .Clear
        .AddItem "By Sale"
        .AddItem "By Scrap"
        .AddItem "«Separated FA"
    End With

    CmdRemove.Caption = "Delete Row"
    Me.Caption = "Disposal of assets"
    Me.Ele.Caption = Me.Caption

    Me.LblShortcutKeys.Caption = "(New F12 OR Enter) ,(Edit F11),(Save F10),(Undo F9),(Delete F8),(Search F7)"
    Me.lbl(4).Caption = "Operation ID"
    Me.lbl(1).Caption = "Operation Date"
    Me.lbl(3).Caption = "Expenses Type"
    Me.lbl(2).Caption = "Total"
    Me.lbl(0).Caption = "Vendor Bill#"
    Me.lbl(5).Caption = "Remarks"
    Me.lbl(8).Caption = "Issued By."
    Me.lbl(7).Caption = "Current Record."
    lbl(27).Caption = "Select Asset"
    Fra.Caption = "GL"
    lbl(11).Caption = "GL#"
    lbl(13).Caption = "interval"
    lbl(9).Caption = "Depit"
    lbl(10).Caption = "Credit"
    lbl(17).Caption = "Bank"
    lbl(18).Caption = "Cheque#"
    lbl(19).Caption = "Due Date"
    lbl(22).Caption = "Vendor"

    Me.Cmd(0).Caption = "&New"
    Me.Cmd(1).Caption = "&Edit"
    Me.Cmd(2).Caption = "&Save"
    Me.Cmd(3).Caption = "&Undo"
    Me.Cmd(4).Caption = "&Delete"
    Me.Cmd(5).Caption = "Sear&ch"
    Me.Cmd(6).Caption = "E&xit"
    Me.Cmd(7).Caption = "&Table View"
    Cmd(8).Caption = "Print"
    Cmd(9).Caption = "Cheque Print"
    Cmd(10).Caption = "GL Print "

    Me.CmdHelp.Caption = "&Help"

    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "Index"
        .TextMatrix(0, .ColIndex("AccountName")) = " Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("des")) = "description"
        .TextMatrix(0, .ColIndex("opr_fullcode")) = "Operation"

    End With

End Sub
