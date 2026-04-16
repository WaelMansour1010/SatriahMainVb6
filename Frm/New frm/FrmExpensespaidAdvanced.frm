VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmExpensespaidAdvanced 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„’—Ê›«  «·„œ›Ê⁄… „ﬁœ„«"
   ClientHeight    =   9540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14250
   Icon            =   "FrmExpensespaidAdvanced.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   14250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15360
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15360
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   5175
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   840
      Width           =   14235
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   1455
         Left            =   120
         TabIndex        =   63
         Top             =   120
         Width           =   14055
         Begin VB.TextBox txtto 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   600
            Width           =   7635
         End
         Begin VB.ComboBox CboPaymentType 
            Height          =   315
            Left            =   8760
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   4095
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   255
            Left            =   5760
            TabIndex        =   3
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmExpensespaidAdvanced.frx":6852
            Height          =   315
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
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
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   8760
            TabIndex        =   2
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   90177537
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10410
            TabIndex        =   79
            Top             =   255
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   285
            Index           =   0
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   720
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   195
            Index           =   15
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   255
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   1
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   285
            Index           =   4
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   240
            Width           =   915
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   1095
         Left            =   120
         TabIndex        =   45
         Top             =   1560
         Width           =   14055
         Begin VB.TextBox TxtSerial 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
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
            Left            =   12360
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   240
            Width           =   945
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   6120
            TabIndex        =   8
            Top             =   240
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCExpensesAdvanced 
            Height          =   315
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCExpenses 
            Height          =   315
            Left            =   6120
            TabIndex        =   9
            Top             =   600
            Width           =   4155
            _ExtentX        =   7329
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCSingle 
            Height          =   315
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   3885
            _ExtentX        =   6853
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lblfd1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·„’—Ê› «·„ﬁœ„"
            Height          =   345
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label lblm 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„›—œ"
            Height          =   300
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   600
            Width           =   1890
         End
         Begin VB.Label lblfp12 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» «·„’—Ê›"
            Height          =   300
            Left            =   10200
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   600
            Width           =   2370
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ  "
            Height          =   315
            Index           =   3
            Left            =   13320
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «”„ «·„’—Ê› «·„ﬁœ„"
            Height          =   285
            Index           =   0
            Left            =   10320
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   240
            Width           =   1890
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   2295
         Left            =   120
         TabIndex        =   44
         Top             =   2640
         Width           =   14055
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2E9E9&
            Height          =   1215
            Left            =   0
            TabIndex        =   80
            Top             =   840
            Width           =   13935
            Begin VB.OptionButton check7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ „Õœœ"
               CausesValidation=   0   'False
               Height          =   315
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   720
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton check6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈œ«—… „Õœœ…"
               CausesValidation=   0   'False
               Height          =   435
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   720
               Width           =   1695
            End
            Begin VB.OptionButton check5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›—⁄ „Õœœ"
               CausesValidation=   0   'False
               Height          =   315
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   240
               Width           =   1455
            End
            Begin VB.OptionButton check4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ÊŸ› „Õœœ"
               CausesValidation=   0   'False
               Height          =   315
               Left            =   11040
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton check3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
               CausesValidation=   0   'False
               Height          =   555
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   120
               Width           =   1215
            End
            Begin MSDataListLib.DataCombo DataCombo1 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   6960
               TabIndex        =   20
               Top             =   240
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   1680
               TabIndex        =   22
               Top             =   240
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo3 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   6960
               TabIndex        =   24
               Top             =   720
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo4 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   1680
               TabIndex        =   26
               Top             =   720
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   810
               Left            =   120
               TabIndex        =   27
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   1429
               ButtonPositionImage=   3
               Caption         =   " ‰›Ì–"
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
               ButtonImage     =   "FrmExpensespaidAdvanced.frx":6867
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Height          =   315
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   480
            Width           =   1695
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   120
            TabIndex        =   74
            Top             =   120
            Width           =   8895
            Begin VB.ComboBox DcbPeriodsID 
               Height          =   315
               ItemData        =   "FrmExpensespaidAdvanced.frx":D0C9
               Left            =   5400
               List            =   "FrmExpensespaidAdvanced.frx":D0CB
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox oldTxtSerial1 
               Alignment       =   2  'Center
               BackColor       =   &H00C0E0FF&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   240
               Width           =   1455
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "FrmExpensespaidAdvanced.frx":D0CD
               Left            =   2880
               List            =   "FrmExpensespaidAdvanced.frx":D0CF
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Text            =   "Combo1"
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·œ›⁄« "
               Height          =   315
               Index           =   8
               Left            =   1680
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
               Height          =   315
               Index           =   7
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   240
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
               Height          =   315
               Index           =   6
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   240
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì»œ√ „‰ :"
               Height          =   315
               Index           =   5
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   240
               Width           =   1110
            End
         End
         Begin VB.OptionButton check2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ÊŸ›Ì‰"
            Height          =   315
            Left            =   10920
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   600
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ”«» „⁄Ì‰"
            Height          =   315
            Left            =   11040
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   255
            Index           =   9
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   240
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Õ—ﬂ… ·‹‹  :"
            Height          =   315
            Index           =   4
            Left            =   12480
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Index           =   1
            Left            =   12720
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   1230
         End
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   0
      Width           =   14385
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   39
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensespaidAdvanced.frx":D0D1
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   38
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensespaidAdvanced.frx":D46B
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   37
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensespaidAdvanced.frx":D805
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   36
         Top             =   240
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   ""
         BackColor       =   16777215
         FontSize        =   12
         FontName        =   "Arial"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmExpensespaidAdvanced.frx":DB9F
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmExpensespaidAdvanced.frx":DF39
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’—Ê›«  «·„œ›Ê⁄… „ﬁœ„«"
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
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   4080
      End
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmExpensespaidAdvanced.frx":FCB8
      Left            =   15240
      List            =   "FrmExpensespaidAdvanced.frx":FCC8
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15600
      TabIndex        =   50
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15240
      TabIndex        =   51
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1545
      Left            =   0
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   8040
      Width           =   14235
      _cx             =   25109
      _cy             =   2725
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   58
         Top             =   600
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   28
            ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":FCE1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   9840
            TabIndex        =   30
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":16543
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11400
            TabIndex        =   29
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   240
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   582
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":168DD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   8040
            TabIndex        =   31
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":1D13F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   1800
            TabIndex        =   34
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   582
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":1D4D9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   240
            TabIndex        =   35
            ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   582
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":1DA73
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   405
            Left            =   6720
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":1DE0D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   3600
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":2466F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   5400
            TabIndex        =   82
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
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
            ButtonImage     =   "FrmExpensespaidAdvanced.frx":24A09
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   120
         TabIndex        =   53
         Top             =   0
         Width           =   3855
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   255
            Width           =   675
         End
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   240
            Width           =   540
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9720
         TabIndex        =   59
         Top             =   120
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   330
         Left            =   7680
         TabIndex        =   84
         ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·’› «·Õ«·Ì"
         BackColor       =   14871017
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
         ButtonImage     =   "FrmExpensespaidAdvanced.frx":2B26B
         ButtonImageDisabled=   "FrmExpensespaidAdvanced.frx":31ACD
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton ISButton4 
         Height          =   330
         Left            =   6000
         TabIndex        =   85
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·ﬂ· "
         BackColor       =   14871017
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
         ButtonImage     =   "FrmExpensespaidAdvanced.frx":50CB7
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   12960
         TabIndex        =   60
         Top             =   120
         Width           =   900
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   1995
      Left            =   0
      TabIndex        =   61
      Top             =   6000
      Width           =   14235
      _cx             =   25109
      _cy             =   3519
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
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777088
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   16
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmExpensespaidAdvanced.frx":57519
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   615
         Left            =   1200
         TabIndex        =   83
         Top             =   960
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15360
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":577AB
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":57B45
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":57EDF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":58279
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":58613
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":589AD
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":58D47
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensespaidAdvanced.frx":592E1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmExpensespaidAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Dim II As Long
Private Sub btnQuery_Click()
Load FrmExpensespaidAdvancedSearch
FrmExpensespaidAdvancedSearch.Show vbModal
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TbExpensespaidAdvanced order by  IDEXP "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetExpensespaidAdvanced Me.DcboBox
    Dcombos.GetAccountingCodes Me.DCExpensesAdvanced
    Dcombos.GetAccountingCodes Me.DCExpenses
    Dcombos.GetYearlyComponents Me.DCSingle
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DataCombo1
    Dcombos.GetBranches Me.DataCombo2
    Dcombos.GetEmpDepartments Me.DataCombo3
    Dcombos.GetYearlyComponents Me.DataCombo4
    Call AdditemTocCmp
    BtnFirst_Click
    ShowTip
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If
   Me.Refresh
ErrTrap:
End Sub
' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Public Sub FiLLRec()
    On Error GoTo ErrTrap
    If TxtModFlg = "E" Then
    If check2.value = True Then
    StrSQL = "Delete From TbExpensespaidJoin Where IDEXP='" & val(TxtSerial1.text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    End If
    RsSavRec.Fields("DateM").value = XPDtbTrans.value
    RsSavRec.Fields("DateH").value = Me.Txt_DateHigri.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    RsSavRec.Fields("PayWay").value = val(Me.CboPaymentType.ListIndex)
    RsSavRec.Fields("Explan").value = IIf(txtto.text <> "", Trim(txtto.text), Null)
    RsSavRec.Fields("ExpIDD").value = IIf(TxtSerial.text <> "", Trim(TxtSerial.text), Null)
    RsSavRec.Fields("ExpName").value = IIf(DcboBox.BoundText <> "", Trim(DcboBox.BoundText), Null)
    RsSavRec.Fields("ExpAcount").value = IIf(DCExpensesAdvanced.BoundText <> "", Trim(DCExpensesAdvanced.BoundText), Null)
    RsSavRec.Fields("ExpAcount1").value = IIf(DCExpenses.BoundText <> "", Trim(DCExpenses.BoundText), Null)
    RsSavRec.Fields("ExpSingle").value = IIf(DCSingle.BoundText <> "", Trim(DCSingle.BoundText), Null)
    
    If check1.value = True Then
    RsSavRec.Fields("EXPCheck").value = 0
    Else
    RsSavRec.Fields("EXPCheck").value = 1
    End If
    
    RsSavRec.Fields("ExpValue").value = IIf(Text1.text <> "", Trim(Text1.text), Null)
    RsSavRec.Fields("ExpMonth").value = IIf(val(DcbPeriodsID.ListIndex) <> -1, val((DcbPeriodsID.ListIndex)), Null)
    RsSavRec.Fields("ExpYear").value = IIf(val(Combo1.ListIndex) <> -1, val(Combo1.ListIndex), Null)
    RsSavRec.Fields("ExpNumber").value = IIf(oldTxtSerial1.text <> "", Trim(oldTxtSerial1.text), Null)
    
    If check3.value = True Then
    RsSavRec.Fields("ExpEmpCheck").value = 1
    End If
    
    If check4.value = True Then
     RsSavRec.Fields("ExpEmpCheck").value = 2
    RsSavRec.Fields("ExpEmpSelect").value = IIf(DataCombo1.BoundText <> "", Trim(DataCombo1.BoundText), Null)
    End If
    
    If check5.value = True Then
    RsSavRec.Fields("ExpEmpCheck").value = 3
    RsSavRec.Fields("ExpBourchSelect").value = IIf(DataCombo2.BoundText <> "", Trim(DataCombo2.BoundText), Null)
    End If
    
    If check6.value = True Then
     RsSavRec.Fields("ExpEmpCheck").value = 4
     RsSavRec.Fields("ExpMangemtSelect").value = IIf(DataCombo3.BoundText <> "", Trim(DataCombo3.BoundText), Null)
     End If
    
    If check7.value = True Then
     RsSavRec.Fields("ExpEmpCheck").value = 5
     RsSavRec.Fields("ExpSingleSelect").value = IIf(DataCombo4.BoundText <> "", Trim(DataCombo4.BoundText), Null)
     End If
    
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ' save grid
     Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TbExpensespaidJoin Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    With Grid
       For i = .FixedRows To .Rows - 1
     If .TextMatrix(i, .ColIndex("Ser")) <> "" Then
                RsDevsub.AddNew
                RsDevsub("IDEXP").value = Me.TxtSerial1.text
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("Emp_ID"))) = "", Null, .TextMatrix(i, .ColIndex("Emp_ID")))
                RsDevsub("BranchID").value = IIf((.TextMatrix(i, .ColIndex("BranchId"))) = "", Null, .TextMatrix(i, .ColIndex("BranchId")))
                RsDevsub("MangmentID").value = IIf((.TextMatrix(i, .ColIndex("DepartmentID"))) = "", Null, .TextMatrix(i, .ColIndex("DepartmentID")))
                RsDevsub("Single").value = IIf((.TextMatrix(i, .ColIndex("mofrad_type"))) = "", Null, .TextMatrix(i, .ColIndex("mofrad_type")))
                RsDevsub("SingleValue").value = IIf((.TextMatrix(i, .ColIndex("Value"))) = "", Null, .TextMatrix(i, .ColIndex("Value")))
                RsDevsub("PayType").value = IIf((.TextMatrix(i, .ColIndex("payment"))) = "", Null, .TextMatrix(i, .ColIndex("payment")))
                RsDevsub("Monthe").value = IIf((.TextMatrix(i, .ColIndex("DcbPeriodsID"))) = "", Null, .TextMatrix(i, .ColIndex("DcbPeriodsID")))
                RsDevsub("SubYear").value = IIf((.TextMatrix(i, .ColIndex("Combo1"))) = "", Null, .TextMatrix(i, .ColIndex("Combo1")))
                RsDevsub("PayValue").value = IIf((.TextMatrix(i, .ColIndex("pymentacount"))) = "", Null, .TextMatrix(i, .ColIndex("pymentacount")))
       RsDevsub.update
      End If
     Next i
     End With
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & Chr(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
               Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            End If
       End Select
  Exit Sub
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If
   End Sub
' full data from database
'+++++++++++++++++++++++++++++++++++++++
Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    Dim i As Integer
    ProgressBar1.Visible = True
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("IDEXP").value), "", RsSavRec.Fields("IDEXP").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("DateM").value), Date, RsSavRec.Fields("DateM").value): ProgressBar1.value = 20
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("DateH").value), "", RsSavRec.Fields("DateH").value): ProgressBar1.value = 30
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 40
    CboPaymentType.ListIndex = IIf(IsNull(RsSavRec.Fields("PayWay").value), "", RsSavRec.Fields("PayWay").value): ProgressBar1.value = 50
    txtto.text = IIf(IsNull(RsSavRec.Fields("Explan").value), "", RsSavRec.Fields("Explan").value): ProgressBar1.value = 60
    TxtSerial.text = IIf(IsNull(RsSavRec.Fields("ExpIDD").value), "", RsSavRec.Fields("ExpIDD").value): ProgressBar1.value = 70
    DcboBox.BoundText = IIf(IsNull(RsSavRec.Fields("ExpName").value), "", RsSavRec.Fields("ExpName").value): ProgressBar1.value = 80
    DCExpenses.BoundText = IIf(IsNull(RsSavRec.Fields("ExpAcount").value), "", RsSavRec.Fields("ExpAcount").value): ProgressBar1.value = 90
    DCExpensesAdvanced.BoundText = IIf(IsNull(RsSavRec.Fields("ExpAcount1").value), "", RsSavRec.Fields("ExpAcount1").value): ProgressBar1.value = 100
    DCSingle.BoundText = IIf(IsNull(RsSavRec.Fields("ExpSingle").value), "", RsSavRec.Fields("ExpSingle").value): ProgressBar1.value = 10
    ''''''''''''''''
     If RsSavRec.Fields("EXPCheck").value = 0 Then
     check1.value = vbChecked
     Me.Grid.Clear flexClearScrollable, flexClearEverything
     Me.Grid.Enabled = False
     Else
     check2.value = vbChecked
     Me.Grid.Clear flexClearScrollable, flexClearEverything
     Me.Grid.Enabled = True
     FillTextGridData
     End If
            ''''''''''''''''''''''''''''''''''''''
     Text1.text = IIf(IsNull(RsSavRec.Fields("ExpValue").value), "", RsSavRec.Fields("ExpValue").value): ProgressBar1.value = 50
     DcbPeriodsID.ListIndex = IIf(IsNull(RsSavRec.Fields("ExpMonth").value), "", RsSavRec.Fields("ExpMonth").value): ProgressBar1.value = 60
     Combo1.ListIndex = IIf(IsNull(RsSavRec.Fields("ExpYear").value), "", RsSavRec.Fields("ExpYear").value)
     oldTxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ExpNumber").value), "", RsSavRec.Fields("ExpNumber").value): ProgressBar1.value = 70
     '''''''''''''''''''''''''''''''''''''''''''''
     If RsSavRec.Fields("ExpEmpCheck").value = 1 Then
     check3.value = vbChecked
     End If
     
     If RsSavRec.Fields("ExpEmpCheck").value = 2 Then
     check4.value = vbChecked
     DataCombo1.BoundText = IIf(IsNull(RsSavRec.Fields("ExpEmpSelect").value), "", RsSavRec.Fields("ExpEmpSelect").value): ProgressBar1.value = 100
     End If
     
     If RsSavRec.Fields("ExpEmpCheck").value = 3 Then
     check5.value = vbChecked
     DataCombo2.BoundText = IIf(IsNull(RsSavRec.Fields("ExpBourchSelect").value), "", RsSavRec.Fields("ExpBourchSelect").value): ProgressBar1.value = 50
     End If
     
     If RsSavRec.Fields("ExpEmpCheck").value = 4 Then
     check6.value = vbChecked
     DataCombo3.BoundText = IIf(IsNull(RsSavRec.Fields("ExpMangemtSelect").value), "", RsSavRec.Fields("ExpMangemtSelect").value): ProgressBar1.value = 40
     End If
     
     If RsSavRec.Fields("ExpEmpCheck").value = 5 Then
     check7.value = vbChecked
     DataCombo4.BoundText = IIf(IsNull(RsSavRec.Fields("ExpSingleSelect").value), "", RsSavRec.Fields("ExpSingleSelect").value)
     End If
     
     DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
     LabCurrRec.Caption = RsSavRec.AbsolutePosition
     LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 80
     ' grid
     
 ProgressBar1.Visible = False
 ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
  Sub FillGridDataWithAdd()
   Dim K As Integer
   Dim H As Integer
   Dim lastrow As Integer
   Dim rs1 As ADODB.Recordset
   Set rs1 = New ADODB.Recordset
   Dim sql As String
   sql = "SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, dbo.TblBranchesData.branch_name, "
   sql = sql & "                    dbo.TblBranchesData.branch_namee, dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.mofrdat.mofrad_name,"
   sql = sql & "                      dbo.EmpSalaryComponent.[value] , dbo.EmpSalaryComponent.AccountCode, dbo.mofrdat.mofrad_namee, dbo.mofrdat.Monthly, dbo.mofrdat.mofrad_type, dbo.mofrdat.mofrad_code"
   sql = sql & "  FROM         dbo.mofrdat RIGHT OUTER JOIN"
   sql = sql & "                      dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode RIGHT OUTER JOIN"
   sql = sql & "                      dbo.TblEmployee ON dbo.EmpSalaryComponent.emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
   sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
   sql = sql & "                      dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
   sql = sql & "  Where (dbo.mofrdat.Monthly = 0)"
   If check4.value = True Then
   sql = sql & " and dbo.TblEmployee.Emp_ID= " & val(DataCombo1.BoundText) & ""
   ElseIf check6.value = True Then
   sql = sql & " and dbo.TblEmployee.DepartmentID= " & val(DataCombo3.BoundText) & ""
    ElseIf check7.value = True Then
   sql = sql & " and dbo.mofrdat.mofrad_code=' " & DataCombo4.BoundText & "'"
   End If
   rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rs1.RecordCount > 0 Then
    With Me.Grid
             lastrow = .Rows
             If rs1.RecordCount > 0 Then
                rs1.MoveFirst
                H = IIf(IsNull(rs1.Fields("Emp_ID").value), 0, rs1.Fields("Emp_ID").value)
               For K = 1 To lastrow - 1
                If val(.TextMatrix(K, .ColIndex("Emp_ID"))) = H Then
             GoTo 10
             End If
             Next K
            .Rows = rs1.RecordCount + lastrow
             Dim i As Integer
             For i = lastrow To rs1.RecordCount + lastrow - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs1.Fields("Emp_ID").value), "", rs1.Fields("Emp_ID").value)
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs1.Fields("Fullcode").value), "", rs1.Fields("Fullcode").value)
                 If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs1.Fields("mofrad_name").value), "", rs1.Fields("mofrad_name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs1.Fields("DepartmentName").value), "", rs1.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs1.Fields("branch_name").value), "", rs1.Fields("branch_name").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs1.Fields("Emp_Name").value), "", rs1.Fields("Emp_Name").value)
                 Else
                .TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs1.Fields("mofrad_namee").value), "", rs1.Fields("mofrad_namee").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs1.Fields("DepartmentNamee").value), "", rs1.Fields("DepartmentNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs1.Fields("branch_namee").value), "", rs1.Fields("branch_namee").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs1.Fields("Emp_Namee").value), "", rs1.Fields("Emp_Namee").value)
                 End If
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs1.Fields("BranchId").value), "", rs1.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("DepartmentID")) = IIf(IsNull(rs1.Fields("DepartmentID").value), "", rs1.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("mofrad_type")) = IIf(IsNull(rs1.Fields("mofrad_code").value), "", rs1.Fields("mofrad_code").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs1.Fields("Value").value), "", rs1.Fields("Value").value)
                .TextMatrix(i, .ColIndex("payment")) = oldTxtSerial1.text
                .TextMatrix(i, .ColIndex("DcbPeriodsID")) = DcbPeriodsID.text
                .TextMatrix(i, .ColIndex("Combo1")) = Combo1.text
                .TextMatrix(i, .ColIndex("pymentacount")) = Round(val(.TextMatrix(i, .ColIndex("Value"))) / val(.TextMatrix(i, .ColIndex("payment"))))
10:              rs1.MoveNext
                 Next
                 rs1.Close
                  End If
                .RowHeight(-1) = 300
        End With
        Else
        MsgBox "⁄›Ê« ....... !!!!!!!! ·« ÌÊÃœ »Ì«‰«  ··≈÷«›…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        End If
 End Sub
  Sub FillTextGridData()
'  If check3.value = True Or check5.value = True Then
    Dim rs1 As ADODB.Recordset
  Set rs1 = New ADODB.Recordset
  Dim sql As String
  sql = "SELECT     dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, "
  sql = sql + "                     dbo.TbExpensespaidJoin.SingleValue, dbo.TbExpensespaidJoin.PayValue, dbo.TbExpensespaidJoin.SubYear, dbo.TbExpensespaidJoin.Monthe, dbo.TbExpensespaidJoin.PayType,"
  sql = sql + "                     dbo.TbExpensespaidJoin.MangmentID, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_ID, dbo.TblBranchesData.branch_id,"
  sql = sql + "                     dbo.TblEmpDepartments.DeparmentID, dbo.TbExpensespaidJoin.Single, dbo.TbExpensespaidJoin.IDEXP, dbo.TbExpensespaidJoin.EmpID, dbo.TbExpensespaidJoin.BranchID,"
  sql = sql + "                     dbo.mofrdat.mofrad_name , dbo.mofrdat.mofrad_namee , dbo.mofrdat.mofrad_code"
  sql = sql + " FROM         dbo.TbExpensespaidJoin LEFT OUTER JOIN"
  sql = sql + "                      dbo.mofrdat ON dbo.TbExpensespaidJoin.Single = dbo.mofrdat.mofrad_code LEFT OUTER JOIN"
  sql = sql + "                      dbo.TblEmpDepartments ON dbo.TbExpensespaidJoin.MangmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
  sql = sql + "                     dbo.TblBranchesData ON dbo.TbExpensespaidJoin.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  sql = sql + "                     dbo.TblEmployee ON dbo.TbExpensespaidJoin.EmpID = dbo.TblEmployee.Emp_ID"
  sql = sql + "  Where (dbo.TbExpensespaidJoin.IDEXP = " & val(TxtSerial1.text) & ") "
    rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If rs1.RecordCount > 0 Then
     rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.Grid
                    For i = .FixedRows To rs1.RecordCount
                   .Rows = .FixedRows + rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs1("Fullcode").value), "", rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs1("Emp_ID").value), 0, rs1("Emp_ID").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs1("Emp_Name").value), "", rs1("Emp_Name").value)
                   .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs1("branch_name").value), "", rs1("branch_name").value)
                   .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs1("DepartmentName").value), "", rs1("DepartmentName").value)
                   .TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs1("mofrad_name").value), "", rs1("mofrad_name").value)
                    Else
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs1("Emp_Namee").value), "", rs1("Emp_Namee").value)
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs1("branch_namee").value), "", rs1("branch_namee").value)
                    .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs1("DepartmentNamee").value), "", rs1("DepartmentNamee").value)
                    .TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs1("mofrad_namee").value), "", rs1("mofrad_namee").value)
                    End If
                   .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs1("BranchID").value), "", rs1("BranchID").value)
                   .TextMatrix(i, .ColIndex("DepartmentID")) = IIf(IsNull(rs1("MangmentID").value), "", rs1("MangmentID").value)
                   .TextMatrix(i, .ColIndex("mofrad_type")) = IIf(IsNull(rs1("mofrad_code").value), "", rs1("mofrad_code").value)
                   .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs1("SingleValue").value), "", rs1("SingleValue").value)
                   .TextMatrix(i, .ColIndex("payment")) = IIf(IsNull(rs1("PayType").value), "", rs1("PayType").value)
                   .TextMatrix(i, .ColIndex("DcbPeriodsID")) = IIf(IsNull(rs1("Monthe").value), "", rs1("Monthe").value)
                   .TextMatrix(i, .ColIndex("Combo1")) = IIf(IsNull(rs1("SubYear").value), "", rs1("SubYear").value)
                   .TextMatrix(i, .ColIndex("pymentacount")) = IIf(IsNull(rs1("PayValue").value), "", rs1("PayValue").value)
                    rs1.MoveNext
             Next i
        End With
        Exit Sub
    '    End If
 End Sub
 Sub FillGridData()
 Dim rs1 As ADODB.Recordset
 Set rs1 = New ADODB.Recordset
 Dim sql As String
 sql = "SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, dbo.TblBranchesData.branch_name, "
 sql = sql & "                    dbo.TblBranchesData.branch_namee, dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.mofrdat.mofrad_name,"
 sql = sql & "                      dbo.EmpSalaryComponent.[value] , dbo.EmpSalaryComponent.AccountCode, dbo.mofrdat.mofrad_namee, dbo.mofrdat.Monthly, dbo.mofrdat.mofrad_type, dbo.mofrdat.mofrad_code"
 sql = sql & "  FROM         dbo.mofrdat RIGHT OUTER JOIN"
 sql = sql & "                      dbo.EmpSalaryComponent ON dbo.mofrdat.mofrad_code = dbo.EmpSalaryComponent.AccountCode RIGHT OUTER JOIN"
 sql = sql & "                      dbo.TblEmployee ON dbo.EmpSalaryComponent.emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 sql = sql & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
 sql = sql & "                      dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
 sql = sql & "  Where (dbo.mofrdat.Monthly = 0)"
 If check4.value = True Then
 sql = sql & " and dbo.TblEmployee.Emp_ID= " & val(DataCombo1.BoundText) & ""
 ElseIf check5.value = True Then
 sql = sql & " and dbo.TblEmployee.BranchId= " & val(DataCombo2.BoundText) & ""
 ElseIf check6.value = True Then
  sql = sql & " and dbo.TblEmployee.DepartmentID= " & val(DataCombo3.BoundText) & ""
 ElseIf check7.value = True Then
  sql = sql & " and dbo.mofrdat.mofrad_code=' " & DataCombo4.BoundText & "'"
 End If
 rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 If rs1.RecordCount > 0 Then
    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable
        If rs1.RecordCount > 0 Then
            .Rows = rs1.RecordCount + 1
            rs1.MoveFirst
            Dim i As Integer
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs1.Fields("Emp_ID").value), "", rs1.Fields("Emp_ID").value)
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs1.Fields("Fullcode").value), "", rs1.Fields("Fullcode").value)
                 If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs1.Fields("mofrad_name").value), "", rs1.Fields("mofrad_name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs1.Fields("DepartmentName").value), "", rs1.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs1.Fields("branch_name").value), "", rs1.Fields("branch_name").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs1.Fields("Emp_Name").value), "", rs1.Fields("Emp_Name").value)
                 Else
                .TextMatrix(i, .ColIndex("mofrad_name")) = IIf(IsNull(rs1.Fields("mofrad_namee").value), "", rs1.Fields("mofrad_namee").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs1.Fields("DepartmentNamee").value), "", rs1.Fields("DepartmentNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs1.Fields("branch_namee").value), "", rs1.Fields("branch_namee").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs1.Fields("Emp_Namee").value), "", rs1.Fields("Emp_Namee").value)
                 End If
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs1.Fields("BranchId").value), "", rs1.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("DepartmentID")) = IIf(IsNull(rs1.Fields("DepartmentID").value), "", rs1.Fields("DepartmentID").value)
                .TextMatrix(i, .ColIndex("mofrad_type")) = IIf(IsNull(rs1.Fields("mofrad_code").value), "", rs1.Fields("mofrad_code").value)
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs1.Fields("Value").value), "", rs1.Fields("Value").value)
                .TextMatrix(i, .ColIndex("payment")) = oldTxtSerial1.text
                .TextMatrix(i, .ColIndex("DcbPeriodsID")) = DcbPeriodsID.text
                .TextMatrix(i, .ColIndex("Combo1")) = Combo1.text
                .TextMatrix(i, .ColIndex("pymentacount")) = Round(val(.TextMatrix(i, .ColIndex("Value"))) / val(.TextMatrix(i, .ColIndex("payment"))))
                 rs1.MoveNext
            Next
            rs1.Close
          End If
        .RowHeight(-1) = 300
    End With
                Else
        MsgBox "⁄›Ê« ....... !!!!!!!! ·« ÌÊÃœ »Ì«‰«  ··≈÷«›…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
 End If
 End Sub

Private Sub ISButton2_Click()
 If DcbPeriodsID.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·‘Â— ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DcbPeriodsID.SetFocus
             Exit Sub
                 Else
            MsgBox "Write Single ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcbPeriodsID.SetFocus
            Exit Sub
            End If
     End If
   '+++++++++++++++++++++++++++++++++++++++++++++++
      If Combo1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Combo1.SetFocus
             Exit Sub
                 Else
            MsgBox "Write Single ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Combo1.SetFocus
            Exit Sub
            End If
     End If
        '+++++++++++++++++++++++++++++++++++++++++++++++
      If oldTxtSerial1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ ⁄œœ «·œ›⁄«  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            oldTxtSerial1.SetFocus
             Exit Sub
                 Else
            MsgBox "Write Single ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            oldTxtSerial1.SetFocus
            Exit Sub
            End If
     End If
  If check4.value = True Then
  If DataCombo1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «”„  «·„ÊŸ› ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo1.SetFocus
             Exit Sub
            DataCombo1.SetFocus
            Else
            MsgBox "Select Employee Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo1.SetFocus
            Exit Sub
            End If
     End If
  End If
  If check5.value = True Then
  If DataCombo2.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «”„  «·›—⁄ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo2.SetFocus
             Exit Sub
            DataCombo2.SetFocus
            Else
            MsgBox "Select Branch Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo2.SetFocus
            Exit Sub
            End If
     End If
  End If
 If check6.value = True Then
  If DataCombo3.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «”„  «·«œ«—… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo3.SetFocus
             Exit Sub
            DataCombo3.SetFocus
            Else
            MsgBox "Select Management Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo3.SetFocus
            Exit Sub
            End If
     End If
  End If
   If check7.value = True Then
  If DataCombo4.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «Œ Ì«— «·„›—œ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DataCombo4.SetFocus
             Exit Sub
            DataCombo4.SetFocus
            Else
            MsgBox "Select Single Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DataCombo4.SetFocus
            Exit Sub
            End If
     End If
  End If
'  IF
    If Me.check4.value = True Or Me.check6.value = True Or Me.check7.value = True Then
    FillGridDataWithAdd
    Else
    Me.Grid.Clear flexClearScrollable, flexClearEverything
    FillGridData
    End If
End Sub
Private Sub ISButton3_Click()
On Error Resume Next
Grid.RemoveItem
End Sub
Private Sub ISButton4_Click()
On Error Resume Next
Me.Grid.Clear flexClearScrollable, flexClearEverything
cleargriid
End Sub
Private Sub Txt_DateHigri_LostFocus()
  VBA.Calendar = vbCalGreg
            XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
End Sub

' change date to hj
  Private Sub XPDtbTrans_Change()
  If Me.TxtModFlg.text <> "R" Then
              Txt_DateHigri.value = ToHijriDate(XPDtbTrans.value)
   End If
   End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Dcbranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Arabic Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
            Dcbranch.SetFocus
         End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
     If CboPaymentType.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· ÿ—Ìﬁ… «·œ›⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            CboPaymentType.SetFocus
            Exit Sub
            Else
            MsgBox "Write Arabic Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
            CboPaymentType.SetFocus
         End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If DcboBox.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ «·„’—Ê› «·„ﬁœ„", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DcboBox.SetFocus
             Exit Sub
             Else
            MsgBox "Write Expenses Advanced ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBox.SetFocus
            Exit Sub
            End If
     End If
    '+++++++++++++++++++++++++++++++++++++++++++++++
       If Text1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·ﬁÌ„… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Text1.SetFocus
             Exit Sub
      Else
            MsgBox "Write Expenses ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Text1.SetFocus
            Exit Sub
            End If
     End If
   '+++++++++++++++++++++++++++++++++++++++++++++++
   If DcbPeriodsID.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·‘Â— ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            DcbPeriodsID.SetFocus
             Exit Sub
                 Else
            MsgBox "Write Single ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcbPeriodsID.SetFocus
            Exit Sub
            End If
     End If
   '+++++++++++++++++++++++++++++++++++++++++++++++
      If Combo1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Combo1.SetFocus
             Exit Sub
                 Else
            MsgBox "Write Single ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Combo1.SetFocus
            Exit Sub
            End If
     End If
        '+++++++++++++++++++++++++++++++++++++++++++++++
      If oldTxtSerial1.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ ⁄œœ «·œ›⁄«  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            oldTxtSerial1.SetFocus
             Exit Sub
                 Else
            MsgBox "Write Single ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            oldTxtSerial1.SetFocus
            Exit Sub
            End If
     End If
     
    ' For Each CtrlTxt In Me.Controls
    '    If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
    '        If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
    '            MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
    '            CtrlTxt.SetFocus
    '            Exit Sub
    '        End If
    '    End If
    'Next
    '------------------------------ check if Empcode exist ----------------------
'   StrVacName = IsRecExist("TbExpensespaidAdvanced", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtSerial1.text) & "'")
  ' If StrVacName <> "" Then
 '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·«”„ „‰ ﬁ»·"
  '     MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
  '    TxtVacName.SetFocus
 '     Exit Sub
'   End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text
            '------------------------------ new record ----------------------------
        Case "N"
                  '------------------------- save record -----------------------------
          AddNewRecored
          AddNewRec
           
        '  BtnLast_Click
        Case "E"
            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select
    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TbExpensespaidAdvanced", "IDEXP", "")
    RsSavRec.AddNew
    RsSavRec.Fields("IDEXP").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "IDEXP=" & RecID, , adSearchForward, 1
    If Not (RsSavRec.EOF) Then
        FiLLTXT
        End If
    Exit Function
ErrTrap:
    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
  End Function
  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub
' undo sub
 Private Sub BtnUndo_Click()
    FindRec val(TxtSerial1.text)
    Me.TxtModFlg.text = "R"
    FiLLTXT
    BtnFirst_Click
End Sub
' refrsh sub
Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click
    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If
    Dim x As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If x = vbNo Then Exit Sub
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                x = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
                RsSavRec.find "IDEXP=" & val(TxtSerial1.text), , adSearchForward, 1
                RsSavRec.Delete
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TbExpensespaidJoin Where IDEXP='" & val(TxtSerial1.text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 If SystemOptions.UserInterface = EnglishInterface Then
                x = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                x = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               End If
               cleargriid
     End If
                            '------------------------------ Move Next ---------------------------.
        Me.Refresh
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
           Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
        Select Case Me.TxtModFlg.text
            Case "N"
                    If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                        End If
                    Case "E"
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
                    Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                 End If
        End Select
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)
        Select Case IntResult
            Case vbYes
               Cancel = True
               btnSave_Click
        Case vbCancel
              Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub EditRec(StrTable As String, _
                   RecID As String)
     FiLLRec
End Sub
'Private Sub Grid_EnterCell()
 '   On Error GoTo ErrTrap
  '  FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("Ser")))
'ErrTrap:
'End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        ISButton1.Enabled = False
     '   Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
       
        
        
    ElseIf TxtModFlg.text = "R" Then
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
    End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        ISButton1.Enabled = True
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
   ElseIf TxtModFlg.text = "E" Then
       Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    '    Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    End If
End Sub

' move btowen recored
Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
If Me.check2.value = True Then
ISButton3.Enabled = True
ISButton4.Enabled = True
Else
ISButton3.Enabled = False
ISButton4.Enabled = False
End If
    RsSavRec.MoveFirst
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
If Me.check2.value = True Then
ISButton3.Enabled = True
ISButton4.Enabled = True
Else
ISButton3.Enabled = False
ISButton4.Enabled = False
End If
    RsSavRec.MoveLast
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.text <> "" Then
        TxtModFlg = "E"
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    cleargriid
    TxtModFlg.text = "N"
    CmbType.ListIndex = 0
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = branch_id
    CmbType.ListIndex = 0
    Dcbranch.SetFocus
    DataCombo1.Enabled = False
    DataCombo2.Enabled = False
    DataCombo3.Enabled = False
    DataCombo4.Enabled = False
    TxtSerial1.Enabled = False
    check1.value = True
    check3.value = True
    Me.Grid.Clear flexClearScrollable, flexClearEverything
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      cleargriid
        Exit Sub
    End If
BegnieWork:
     If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext
        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        cleargriid
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
     cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = Chr(13) + Chr(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With
ErrTrap:
End Sub
' short cut for keys
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If
    'End If
    Exit Sub
ErrTrap:
End Sub
' print Events
'++++++++++++++++++++++++++++++++++++++++++
Private Sub BtnPrint_Click()
On Error GoTo ErrTrap
  If val(Me.TxtSerial1.text) <> 0 Then
      print_report
  End If
ErrTrap:
End Sub
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
   If val(Me.TxtSerial1.text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    sql = "SELECT     dbo.TbExpensespaidAdvanced.IDEXP, dbo.TbExpensespaidAdvanced.DateM, dbo.TbExpensespaidAdvanced.DateH, dbo.TbExpensespaidAdvanced.BranchID, dbo.TblBranchesData.branch_name, "
  sql = sql & "                    dbo.TblBranchesData.branch_namee, dbo.TbExpensespaidAdvanced.PayWay, dbo.TbExpensespaidAdvanced.Explan, dbo.TbExpensespaidAdvanced.ExpIDD,"
  sql = sql & "                    dbo.TbExpensespaidAdvanced.ExpName, dbo.TbExpensesprovided.name, dbo.TbExpensesprovided.namee, dbo.TbExpensespaidAdvanced.ExpAcount, dbo.ACCOUNTS.Account_Name,"
  sql = sql & "                    dbo.ACCOUNTS.Account_NameEng, dbo.TbExpensespaidAdvanced.ExpAcount1, ACCOUNTS_1.Account_Name AS Account_Name1, ACCOUNTS_1.Account_NameEng AS Account_Name1E,"
  sql = sql & "                    dbo.TbExpensespaidAdvanced.ExpSingle, dbo.mofrdat.mofrad_name, dbo.mofrdat.mofrad_namee, dbo.TbExpensespaidAdvanced.EXPCheck, dbo.TbExpensespaidAdvanced.ExpValue,"
  sql = sql & "                    dbo.TbExpensespaidAdvanced.ExpMonth, dbo.TbExpensespaidAdvanced.ExpYear, dbo.TbExpensespaidAdvanced.ExpNumber, dbo.TbExpensespaidAdvanced.ExpEmpCheck,"
  sql = sql & "                    dbo.TbExpensespaidAdvanced.ExpEmpSelect, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
 sql = sql & "                     dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1,"
  sql = sql & "                    dbo.TblEmployee.Emp_Namee, dbo.TbExpensespaidAdvanced.ExpBourchSelect, TblBranchesData_1.branch_name AS branch_nameSelect,"
 sql = sql & "                     TblBranchesData_1.branch_namee AS branch_nameSelectE, dbo.TbExpensespaidAdvanced.ExpMangemtSelect, dbo.TblEmpDepartments.DepartmentName,"
 sql = sql & "                     dbo.TblEmpDepartments.DepartmentNamee, dbo.TbExpensespaidAdvanced.ExpSingleSelect, mofrdat_1.mofrad_name AS mofrad_nameSelct, mofrdat_1.mofrad_namee AS mofrad_nameSelctE,"
 sql = sql & "                     dbo.TbExpensespaidJoin.EmpID, TblEmployee_1.Emp_Name AS Emp_NameDet, TblEmployee_1.Emp_Name1 AS Emp_NameDet1, TblEmployee_1.Emp_Name2 AS Emp_NameDet2,"
 sql = sql & "                     TblEmployee_1.Emp_Name3 AS Emp_NameDet3, TblEmployee_1.Emp_Name4 AS Emp_NameDet4, TblEmployee_1.Fullcode AS FullcodeDet, TblEmployee_1.Emp_Namee4 AS Emp_NameeDet4,"
 sql = sql & "                     TblEmployee_1.Emp_Namee3 AS Emp_NameeDet3, TblEmployee_1.Emp_Namee2 AS Emp_NameeDet2, TblEmployee_1.Emp_Namee1 AS Emp_NameeDet1,"
 sql = sql & "                     TblEmployee_1.Emp_Namee AS Emp_NameeDet, dbo.TbExpensespaidJoin.BranchID AS BranchIDDet, TblBranchesData_2.branch_name AS branch_nameDet,"
 sql = sql & "                     TblBranchesData_2.branch_namee AS branch_nameDetE, dbo.TbExpensespaidJoin.MangmentID, TblEmpDepartments_1.DepartmentName AS DepartmentNameDet,"
 sql = sql & "                     TblEmpDepartments_1.DepartmentNamee AS DepartmentNameeDet, dbo.TbExpensespaidJoin.Single, mofrdat_2.mofrad_name AS mofrad_nameDet, mofrdat_2.mofrad_namee AS mofrad_nameDetE,"
sql = sql & "                       dbo.TbExpensespaidJoin.SingleValue , dbo.TbExpensespaidJoin.PayType, dbo.TbExpensespaidJoin.Monthe, dbo.TbExpensespaidJoin.SubYear, dbo.TbExpensespaidJoin.PayValue , dbo.TbExpensespaidJoin.ID "
sql = sql & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
 sql = sql & "                     dbo.TbExpensespaidAdvanced LEFT OUTER JOIN"
 sql = sql & "                     dbo.mofrdat mofrdat_2 RIGHT OUTER JOIN"
 sql = sql & "                     dbo.TbExpensespaidJoin ON mofrdat_2.mofrad_code = dbo.TbExpensespaidJoin.Single LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblEmpDepartments TblEmpDepartments_1 ON dbo.TbExpensespaidJoin.MangmentID = TblEmpDepartments_1.DeparmentID LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblBranchesData TblBranchesData_2 ON dbo.TbExpensespaidJoin.BranchID = TblBranchesData_2.branch_id LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TbExpensespaidJoin.EmpID = TblEmployee_1.Emp_ID ON dbo.TbExpensespaidAdvanced.IDEXP = dbo.TbExpensespaidJoin.IDEXP ON"
 sql = sql & "                     dbo.TblEmployee.Emp_ID = dbo.TbExpensespaidAdvanced.ExpEmpSelect LEFT OUTER JOIN"
 sql = sql & "                     dbo.mofrdat mofrdat_1 ON dbo.TbExpensespaidAdvanced.ExpSingleSelect = mofrdat_1.mofrad_code LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblEmpDepartments ON dbo.TbExpensespaidAdvanced.ExpMangemtSelect = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
 sql = sql & "                     dbo.TblBranchesData TblBranchesData_1 ON dbo.TbExpensespaidAdvanced.ExpBourchSelect = TblBranchesData_1.branch_id LEFT OUTER JOIN"
 sql = sql & "                     dbo.mofrdat ON dbo.TbExpensespaidAdvanced.ExpSingle = dbo.mofrdat.mofrad_code LEFT OUTER JOIN"
 sql = sql & "                     dbo.ACCOUNTS ACCOUNTS_1 ON dbo.TbExpensespaidAdvanced.ExpAcount1 = ACCOUNTS_1.Account_Code LEFT OUTER JOIN"
 sql = sql & "                     dbo.ACCOUNTS ON dbo.TbExpensespaidAdvanced.ExpAcount = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
 sql = sql & "                     dbo.TbExpensesprovided ON dbo.TbExpensespaidAdvanced.ExpName = dbo.TbExpensesprovided.ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.TbExpensespaidAdvanced.BranchID = dbo.TblBranchesData.branch_id"
sql = sql & " Where (dbo.TbExpensespaidAdvanced.IDEXP = " & val(TxtSerial1.text) & ")"
                    

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ExpensespaidAdvancedRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ExpensespaidAdvancedRPTEE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
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
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function
' chang langeg Event
'++++++++++++++++++++++++++++++++++++
'Private Sub TxtVacName_GotFocus()
 '   SwitchKeyboardLang LANG_ARABIC
'End Sub
'Private Sub TxtVacNamee_GotFocus()
'SwitchKeyboardLang LANG_ENGLISH
'End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Expenses Paid Advanced"
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Operation ID"
    Me.lbl(2).Caption = "Date"
    Me.lbl(1).Caption = "HJ Date"
    Me.Label3.Caption = "Branch"
    Me.lbl(15).Caption = "Payment Method"
    Me.lbl(0).Caption = "based on "
    '''''''''''''' next
    Me.Label1(3).Caption = "Code"
    Me.Label1(0).Caption = "Submitted Expense Name"
    Me.lblfd1.Caption = "Submitted Expense Account"
    Me.lblfp12.Caption = "Expense Account"
    Me.lblm.Caption = "Single"
    '''''''''''''''''''''''' next
    Me.Label1(4).Caption = "Movement for"
    Me.check1.Caption = "Select Acount"
    Me.check2.Caption = "Employees"
    Me.Label1(9).Caption = "Value"
    Me.Label1(5).Caption = "Start From"
    Me.Label1(6).Caption = "Month"
    Me.Label1(7).Caption = "Year"
    Me.Label1(8).Caption = "Payments NO."
    Me.check3.Caption = "All Employees"
    Me.check4.Caption = "Selsct Employee"
    Me.check5.Caption = "Selsct Branch"
    Me.check6.Caption = "Selsct Department"
    Me.check7.Caption = "Selsct Single"
    Me.ISButton2.Caption = "OK"
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    ISButton3.Caption = "Delet Select"
    ISButton4.Caption = "Delet All"
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Employee Code"
        .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch Name"
        .TextMatrix(0, .ColIndex("DepartmentName")) = "Department Name"
        .TextMatrix(0, .ColIndex("mofrad_name")) = "Single Name"
        .TextMatrix(0, .ColIndex("Value")) = "Single Value"
        .TextMatrix(0, .ColIndex("payment")) = "Spread Over Payments Number"
        .TextMatrix(0, .ColIndex("DcbPeriodsID")) = "Started Monthe"
        .TextMatrix(0, .ColIndex("Combo1")) = "Started Year"
        .TextMatrix(0, .ColIndex("pymentacount")) = "Installment Value"
      End With
ErrTrap:
End Sub
 Private Sub DcboBox_Change()
  DcboBox_Click (0)
 End Sub
 Private Sub DcboBox_Click(Area As Integer)
 If Me.TxtModFlg <> "R" Then
 If val(DcboBox.BoundText) <> 0 Then
 RetriveData val(DcboBox.BoundText)
 End If
 End If
 End Sub
 Sub RetriveData(Optional id As Integer = 0)
 On Error GoTo ErrTrap
 Dim rs1 As ADODB.Recordset
 Set rs1 = New ADODB.Recordset
 Dim sql As String
 sql = "select * from TbExpensesprovided where ID= " & id & ""
 rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 If rs1.RecordCount > 0 Then
 TxtSerial.text = IIf(IsNull(rs1("ID").value), "", rs1("ID").value)
 DCExpensesAdvanced.BoundText = IIf(IsNull(rs1("Exprovided").value), "", rs1("Exprovided").value)
 DCExpenses.BoundText = IIf(IsNull(rs1("Expenses").value), "", rs1("Expenses").value)
 DCSingle.BoundText = IIf(IsNull(rs1("Single").value), "", rs1("Single").value)
 End If
ErrTrap:
 End Sub
' check box event
'+++++++++++++++++++++++++++++++++++
  Private Sub Check1_Click()
  On Error GoTo ErrTrap
  If check1.value = vbChecked Then
  Text1.Enabled = False
   Else
  Text1.text = ""
  Text1.Enabled = True
  Grid.Enabled = False
  check3.Enabled = False
  check4.Enabled = False
  check5.Enabled = False
  check6.Enabled = False
  check7.Enabled = False
   DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  ISButton2.Enabled = False
  ISButton4.Enabled = False
  ISButton3.Enabled = False
  Me.Grid.Clear flexClearScrollable, flexClearEverything
  End If
ErrTrap:
 End Sub
 Private Sub Check2_Click()
 On Error GoTo ErrTrap
 If check2.value = vbChecked Then
  Text1.Enabled = True
  Else
  Text1.text = "0"
  Text1.Enabled = False
   Grid.Enabled = True
    check3.Enabled = True
  check4.Enabled = True
  check5.Enabled = True
  check6.Enabled = True
  check7.Enabled = True
  ISButton3.Enabled = True
  ISButton4.Enabled = True
  ISButton2.Enabled = True
  Me.Grid.Clear flexClearScrollable, flexClearEverything
  End If
ErrTrap:
End Sub
 Private Sub Check3_Click()
 On Error GoTo ErrTrap
  DataCombo1.Enabled = False
  DataCombo1.text = ""
  DataCombo2.Enabled = False
  DataCombo2.text = ""
  DataCombo3.Enabled = False
  DataCombo3.text = ""
  DataCombo4.Enabled = False
  DataCombo4.text = ""
ErrTrap:
End Sub
Private Sub check4_Click()
 On Error GoTo ErrTrap
If check4.value = vbChecked Then
  DataCombo1.Enabled = False
  Else
  DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo1.Enabled = True
  DataCombo2.Enabled = False
  DataCombo3.Enabled = False
  DataCombo4.Enabled = False
  End If
ErrTrap:
End Sub
Private Sub check5_Click()
 On Error GoTo ErrTrap
If check5.value = vbChecked Then
  DataCombo2.Enabled = False
  Else
   DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo2.Enabled = True
  DataCombo1.Enabled = False
  DataCombo3.Enabled = False
  DataCombo4.Enabled = False
  End If
ErrTrap:
End Sub
Private Sub check6_Click()
 On Error GoTo ErrTrap
If check6.value = vbChecked Then
  DataCombo3.Enabled = False
  Else
  DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo3.Enabled = True
  DataCombo1.Enabled = False
  DataCombo2.Enabled = False
  DataCombo4.Enabled = False
  End If
ErrTrap:
End Sub
Private Sub check7_Click()
 On Error GoTo ErrTrap
If check7.value = vbChecked Then
  DataCombo4.Enabled = False
  Else
  DataCombo1.text = ""
  DataCombo2.text = ""
  DataCombo3.text = ""
  DataCombo4.text = ""
  DataCombo4.Enabled = True
  DataCombo1.Enabled = False
  DataCombo2.Enabled = False
  DataCombo3.Enabled = False
  End If
ErrTrap:
End Sub
' key press
'+++++++++++++++++++++++++++++++++++++++++++++
'Private Sub TxtVacName_KeyPress(KeyAscii As Integer)
'   On Error GoTo ErrTrap
'  If KeyAscii = 13 Then
'  TxtVacNamee.SetFocus
'  End If
'ErrTrap:
'End Sub
'Private Sub TxtVacNamee_KeyPress(KeyAscii As Integer)
'   On Error GoTo ErrTrap
'  If KeyAscii = 13 Then
'  DCExpensesAdvanced.SetFocus
'  End If
'ErrTrap:
'End Sub
Private Sub DCExpensesAdvanced_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  DCExpenses.SetFocus
  End If
ErrTrap:
End Sub
Private Sub DCExpenses_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  DCSingle.SetFocus
  End If
ErrTrap:
End Sub
Private Sub DCSingle_KeyPress(KeyAscii As Integer)
   On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  Call btnSave_Click
  End If
ErrTrap:
End Sub
Private Sub AdditemTocCmp()
 On Error GoTo ErrTrap
   Dim i As Integer
  ' full cop month
  DcbPeriodsID.Clear
    For i = 1 To 12
    DcbPeriodsID.AddItem i
    Next i
    Combo1.Clear
    'full cop year
    For i = 2014 To 2050
    Combo1.AddItem i
    Next
    ' full pay way
   If SystemOptions.UserInterface = EnglishInterface Then
    With Me.CboPaymentType
        .Clear
        .AddItem "Cassh /custody"
        .AddItem "Check"
        .AddItem "Bank Transfer"
        .AddItem "Check Outstanding"
        .AddItem "Acount"
        .AddItem "Bank Ordered"
    End With
    Else
    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ/ ⁄ÂœÂ"
        .AddItem "‘Ìﬂ"
        .AddItem " ÕÊ«·Â »‰ﬂÌÂ"
        .AddItem "‘Ìﬂ „”œœ"
        .AddItem "Õ”«»"
        .AddItem "√„— »‰ﬂÌ"
    End With
    End If
ErrTrap:
End Sub
Private Sub cleargriid()
Me.Grid.Rows = 1
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TbExpensespaidAdvanced"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end




