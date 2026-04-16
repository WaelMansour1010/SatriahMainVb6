VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmOpenClosPeriod 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   FillColor       =   &H80000012&
   Icon            =   "FrmOpenClosPeriod.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   14235
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmOpenClosPeriod.frx":6852
      Left            =   15480
      List            =   "FrmOpenClosPeriod.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   14505
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   19
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
         ButtonImage     =   "FrmOpenClosPeriod.frx":687B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   20
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
         ButtonImage     =   "FrmOpenClosPeriod.frx":6C15
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   21
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
         ButtonImage     =   "FrmOpenClosPeriod.frx":6FAF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   22
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
         ButtonImage     =   "FrmOpenClosPeriod.frx":7349
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "    › Õ ≈ﬁ›«· «·› —«      "
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
         TabIndex        =   23
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmOpenClosPeriod.frx":76E3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   14235
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   14055
         Begin VB.ComboBox DcbType 
            DataSource      =   "Adodc1"
            Height          =   315
            ItemData        =   "FrmOpenClosPeriod.frx":8AE8
            Left            =   6000
            List            =   "FrmOpenClosPeriod.frx":8AEA
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   1635
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   9000
            TabIndex        =   14
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   98697217
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmOpenClosPeriod.frx":8AEC
            Height          =   315
            Left            =   2640
            TabIndex        =   43
            Top             =   240
            Width           =   2775
            _ExtentX        =   4895
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
         Begin MSDataListLib.DataCombo DcbYear 
            Bindings        =   "FrmOpenClosPeriod.frx":8B01
            Height          =   315
            Left            =   240
            TabIndex        =   79
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            Caption         =   "«·”‰…"
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   80
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·≈Ã—«¡"
            Height          =   285
            Index           =   12
            Left            =   7680
            TabIndex        =   46
            Top             =   255
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   4920
            TabIndex        =   44
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10410
            TabIndex        =   15
            Top             =   255
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "›Ì Õ«·… «·› Õ"
         ForeColor       =   &H000040C0&
         Height          =   7455
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   720
         Width           =   14055
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕœÌœ «·ﬂ·"
            Height          =   195
            Left            =   12480
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox TxtSearchCode2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11310
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   240
            Width           =   1065
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  «·› Õ"
            ForeColor       =   &H000040C0&
            Height          =   3735
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   3240
            Width           =   13935
            Begin VB.ListBox ListGroupSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1425
               ItemData        =   "FrmOpenClosPeriod.frx":8B16
               Left            =   120
               List            =   "FrmOpenClosPeriod.frx":8B18
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   2160
               Width           =   5175
            End
            Begin VB.ListBox ListGroup 
               Height          =   1425
               ItemData        =   "FrmOpenClosPeriod.frx":8B1A
               Left            =   5760
               List            =   "FrmOpenClosPeriod.frx":8B21
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   2160
               Width           =   5775
            End
            Begin VB.ListBox ListUserAll 
               Height          =   1425
               ItemData        =   "FrmOpenClosPeriod.frx":8B33
               Left            =   5760
               List            =   "FrmOpenClosPeriod.frx":8B3A
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   600
               Width           =   5775
            End
            Begin VB.ListBox ListUserAllSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1425
               ItemData        =   "FrmOpenClosPeriod.frx":8B4B
               Left            =   120
               List            =   "FrmOpenClosPeriod.frx":8B4D
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   600
               Width           =   5175
            End
            Begin XtremeSuiteControls.CheckBox Ch 
               Height          =   255
               Index           =   8
               Left            =   12000
               TabIndex        =   52
               Top             =   240
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Ì„ﬂ‰ «· ⁄œÌ·"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox Ch 
               Height          =   255
               Index           =   9
               Left            =   9960
               TabIndex        =   53
               Top             =   240
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Ì„ﬂ‰ «·ÿ»«⁄…"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox Ch 
               Height          =   255
               Index           =   10
               Left            =   8280
               TabIndex        =   54
               Top             =   240
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Ì„ﬂ‰ «·Õ–›"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox Ch 
               Height          =   255
               Index           =   11
               Left            =   6600
               TabIndex        =   55
               Top             =   240
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Ì„ﬂ‰ «·«÷«›…"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   255
               Index           =   0
               Left            =   12000
               TabIndex        =   56
               Top             =   960
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "·„” Œœ„ „⁄Ì‰"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   255
               Index           =   1
               Left            =   12000
               TabIndex        =   65
               Top             =   2640
               Width           =   1575
               _Version        =   786432
               _ExtentX        =   2778
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "·„Ã„Ê⁄…"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Opt 
               Height          =   255
               Index           =   2
               Left            =   4800
               TabIndex        =   66
               Top             =   240
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "·Ã„Ì⁄ «·„” Œœ„Ì‰"
               ForeColor       =   16711680
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   2280
               Width           =   495
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   2520
               Width           =   495
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   2760
               Width           =   495
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   3000
               Width           =   495
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   1440
               Width           =   495
            End
         End
         Begin MSDataListLib.DataCombo DcboEmpName2 
            Bindings        =   "FrmOpenClosPeriod.frx":8B4F
            Height          =   315
            Left            =   6000
            TabIndex        =   58
            Top             =   240
            Width           =   5295
            _ExtentX        =   9340
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   2175
            Left            =   120
            TabIndex        =   63
            Top             =   960
            Width           =   13875
            _cx             =   24474
            _cy             =   3836
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmOpenClosPeriod.frx":8B64
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
            Begin MSComctlLib.ProgressBar ProgressBar2 
               Height          =   615
               Left            =   1200
               TabIndex        =   64
               Top             =   -6600
               Visible         =   0   'False
               Width           =   11295
               _ExtentX        =   19923
               _ExtentY        =   1085
               _Version        =   393216
               Appearance      =   0
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —«  «·„ﬁ›·…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   405
            Index           =   3
            Left            =   5160
            TabIndex        =   62
            Top             =   600
            Width           =   2565
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ«∆„ »«·› Õ"
            Height          =   285
            Index           =   0
            Left            =   12360
            TabIndex        =   59
            Top             =   240
            Width           =   1605
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "›Ì Õ«·… «·≈ﬁ›«·"
         ForeColor       =   &H000040C0&
         Height          =   6615
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   14055
         Begin VB.CheckBox ChkAll 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕœÌœ «·ﬂ·"
            Height          =   195
            Left            =   12600
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   720
            Width           =   1095
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Height          =   2295
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   4200
            Width           =   8775
            Begin VB.Frame Frame9 
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1815
               Left            =   2160
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   120
               Width           =   2655
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   4
                  Left            =   120
                  TabIndex        =   97
                  Top             =   240
                  Width           =   2415
                  _Version        =   786432
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„” ‰œ«  ﬁÌœ «·«⁄ „«œ"
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   98
                  Top             =   600
                  Width           =   2415
                  _Version        =   786432
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   " ”ÊÌ… «·»‰Êﬂ"
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   99
                  Top             =   960
                  Width           =   2415
                  _Version        =   786432
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„‘ —Ì«  ·„  ”·„"
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   100
                  Top             =   1320
                  Width           =   2415
                  _Version        =   786432
                  _ExtentX        =   4260
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„»Ì⁄«  ·„  ”·„"
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   1815
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   120
               Width           =   2655
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   0
                  Left            =   360
                  TabIndex        =   92
                  Top             =   240
                  Width           =   2175
                  _Version        =   786432
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "√ﬁ”«ÿ «·√’Ê· «·À«» Â"
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   1
                  Left            =   360
                  TabIndex        =   93
                  Top             =   600
                  Width           =   2175
                  _Version        =   786432
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ﬁÌÊœ ≈” Õﬁ«ﬁ «·—« »"
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   2
                  Left            =   360
                  TabIndex        =   94
                  Top             =   960
                  Width           =   2175
                  _Version        =   786432
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«ÿ›«¡ «·„ﬁœ„« "
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.CheckBox Ch 
                  Height          =   255
                  Index           =   3
                  Left            =   360
                  TabIndex        =   95
                  Top             =   1320
                  Width           =   2175
                  _Version        =   786432
                  _ExtentX        =   3836
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«‰‘«¡ «·„Œ’’« "
                  ForeColor       =   16711680
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   0
               Left            =   4200
               TabIndex        =   83
               Top             =   360
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   1
               Left            =   4200
               TabIndex        =   84
               Top             =   720
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   2
               Left            =   4200
               TabIndex        =   85
               Top             =   1080
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   3
               Left            =   4200
               TabIndex        =   86
               Top             =   1440
               Width           =   1695
               _Version        =   786432
               _ExtentX        =   2990
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   87
               Top             =   360
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   88
               Top             =   720
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   89
               Top             =   1080
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChClose 
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   90
               Top             =   1440
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "≈ﬁ›«· ÌœÊÌ"
               ForeColor       =   0
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.CheckBox ChAll 
               Height          =   255
               Left            =   600
               TabIndex        =   102
               Top             =   1800
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«ﬁ›«·  ÌœÊÌ ··ﬂ·"
               ForeColor       =   16512
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11070
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   1305
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Bindings        =   "FrmOpenClosPeriod.frx":8C39
            Height          =   315
            Left            =   6000
            TabIndex        =   1
            Top             =   240
            Width           =   5055
            _ExtentX        =   8916
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3135
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   13875
            _cx             =   24474
            _cy             =   5530
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmOpenClosPeriod.frx":8C4E
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
               TabIndex        =   48
               Top             =   -5640
               Visible         =   0   'False
               Width           =   11295
               _ExtentX        =   19923
               _ExtentY        =   1085
               _Version        =   393216
               Appearance      =   0
            End
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   615
            Left            =   120
            TabIndex        =   82
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   4320
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   1085
            Caption         =   " ‰›Ì– «·«ﬁ›«· "
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":8D23
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " ‰›Ì– «·«ﬁ›«· «·ÌœÊÌ »⁄œ   ‰›Ì– «·«ﬁ››«· «·«·Ì"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   525
            Index           =   6
            Left            =   -120
            TabIndex        =   101
            Top             =   5760
            Width           =   4485
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —«  «·„› ÊÕ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   405
            Index           =   1
            Left            =   4920
            TabIndex        =   61
            Top             =   600
            Width           =   2805
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁ«∆„ »«·≈ﬁ›«·"
            Height          =   285
            Index           =   10
            Left            =   12360
            TabIndex        =   42
            Top             =   240
            Width           =   1605
         End
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   26
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
      Left            =   15480
      TabIndex        =   27
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
      Height          =   2985
      Left            =   0
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6720
      Width           =   14235
      _cx             =   25109
      _cy             =   5265
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   30
         Top             =   1560
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   255
            Width           =   675
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   1
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   29
         Top             =   2160
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   3
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":F585
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   5
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":15DE7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   4
            ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":16181
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   6
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":1C9E3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   7
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Visible         =   0   'False
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":1CD7D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   8
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":1D317
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1560
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   582
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
            ButtonImage     =   "FrmOpenClosPeriod.frx":1D6B1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9000
         TabIndex        =   35
         Top             =   1680
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   12840
         TabIndex        =   36
         Top             =   1680
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
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
            Picture         =   "FrmOpenClosPeriod.frx":1DA4B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOpenClosPeriod.frx":1DDE5
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOpenClosPeriod.frx":1E17F
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOpenClosPeriod.frx":1E519
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOpenClosPeriod.frx":1E8B3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOpenClosPeriod.frx":1EC4D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOpenClosPeriod.frx":1EFE7
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOpenClosPeriod.frx":1F581
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
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
      ButtonImage     =   "FrmOpenClosPeriod.frx":1F91B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
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
      ButtonImage     =   "FrmOpenClosPeriod.frx":2617D
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
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
      ButtonImage     =   "FrmOpenClosPeriod.frx":2C9DF
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
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
      Left            =   15480
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmOpenClosPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim II As Long
 Public LonRow As Double
Public LngCol As Double
Sub FilGrid12(Optional ID As Double = 0)
Dim RsDev As ADODB.Recordset
Dim i As Integer
Dim StrSQL As String
    StrSQL = " SELECT   * FROM         dbo.TblAccountIntervals  "
    StrSQL = StrSQL & "  where TblyearsdataId=" & ID & " and OpenState=1"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
      
                If RsDev("OpenState").value = 1 Then
                    .Cell(flexcpChecked, i, .ColIndex("OpenState")) = flexChecked
                    
                Else
                    .Cell(flexcpChecked, i, .ColIndex("OpenState")) = flexUnchecked
                End If
            
                .TextMatrix(i, .ColIndex("StartDate")) = IIf(IsNull(RsDev("StartDate").value), Date, (RsDev("StartDate").value))
                '.TextMatrix(i, .ColIndex("AccountIntervalID")) = IIf(IsNull(RsDev("AccountIntervalID").value), "", (RsDev("AccountIntervalID").value))
            
                .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(RsDev("EndDate").value), Date, (RsDev("EndDate").value))
  
                .TextMatrix(i, .ColIndex("Comment")) = IIf(IsNull(RsDev("Comment").value), "", RsDev("Comment").value)
            .AutoSize 0, .Cols - 1, False
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
   ' ReLineGrid
End Sub
Sub filgrid(Optional ID As Double = 0)
Dim RsDev As ADODB.Recordset
Dim i As Integer
Dim StrSQL As String
    StrSQL = " SELECT   * FROM         dbo.TblAccountIntervals  "
    StrSQL = StrSQL & "  where TblyearsdataId=" & ID & " and OpenState<>1"
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
      
                If RsDev("OpenState").value = 1 Then
                    .Cell(flexcpChecked, i, .ColIndex("OpenState")) = flexChecked
                    
                Else
                    .Cell(flexcpChecked, i, .ColIndex("OpenState")) = flexUnchecked
                End If
            
                .TextMatrix(i, .ColIndex("StartDate")) = IIf(IsNull(RsDev("StartDate").value), Date, (RsDev("StartDate").value))
                '.TextMatrix(i, .ColIndex("AccountIntervalID")) = IIf(IsNull(RsDev("AccountIntervalID").value), "", (RsDev("AccountIntervalID").value))
            
                .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(RsDev("EndDate").value), Date, (RsDev("EndDate").value))
  
                .TextMatrix(i, .ColIndex("Comment")) = IIf(IsNull(RsDev("Comment").value), "", RsDev("Comment").value)
            .AutoSize 0, .Cols - 1, False
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
   ' ReLineGrid
End Sub

Function CheCkSalary(Optional sgn1 As String) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
CheCkSalary = False
sql = " SELECT    sgn"
sql = sql & " From dbo.emp_salary"
sql = sql & "  Where sgn='" & sgn1 & "'"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheCkSalary = True
Else
CheCkSalary = False
End If
End Function
Function CheCkFixedAssest(FromTransDate As Date, TOtransDate As Date) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
CheCkFixedAssest = False
sql = " SELECT    RecordDate"
sql = sql & " From dbo.FixedAssetInstallments"
sql = sql & "  Where( RecordDate >=" & SQLDate(FromTransDate, True) & ") And (RecordDate <= " & SQLDate(TOtransDate, True) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheCkFixedAssest = True
Else
CheCkFixedAssest = False
End If
End Function
Function CheckApprove(FromTransDate As Date, TOtransDate As Date) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
CheckApprove = False
sql = " SELECT SendTime  "
sql = sql & " From dbo.ApprovalData"
sql = sql & "  Where( SendTime >=" & SQLDate(FromTransDate, True) & ") And (SendTime <= " & SQLDate(TOtransDate, True) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckApprove = True
Else
CheckApprove = False
End If
End Function

Function CheckPurche(FromTransDate As Date, TOtransDate As Date) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
CheckPurche = False
sql = "SELECT     Nots, Transaction_Date, Transaction_Type"
sql = sql & " From dbo.Transactions"
sql = sql & "  WHERE     (Nots = N' ' OR"
sql = sql & "                       Nots IS NULL) AND (Transaction_Type = 22)"
sql = sql & "  and( Transaction_Date >=" & SQLDate(FromTransDate, True) & ") And (Transaction_Date <= " & SQLDate(TOtransDate, True) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckPurche = True
Else
CheckPurche = False
End If
End Function
Function CheckSales(FromTransDate As Date, TOtransDate As Date) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
CheckSales = False
sql = "SELECT     Nots, Transaction_Date, Transaction_Type"
sql = sql & " From dbo.Transactions"
sql = sql & "  WHERE     (Nots = N' ' OR"
sql = sql & "                       Nots IS NULL) AND (Transaction_Type = 21)"
sql = sql & "  and( Transaction_Date >=" & SQLDate(FromTransDate, True) & ") And (Transaction_Date <= " & SQLDate(TOtransDate, True) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckSales = True
Else
CheckSales = False
End If
End Function
Function CheckLBankSettlement(FromTransDate As Date, TOtransDate As Date) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
CheckLBankSettlement = False
sql = " SELECT SettlementDT  "
sql = sql & " From dbo.TBLBankSettlement"
sql = sql & "  Where( SettlementDT >=" & SQLDate(FromTransDate, True) & ") And (SettlementDT <= " & SQLDate(TOtransDate, True) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckLBankSettlement = True
Else
CheckLBankSettlement = False
End If
End Function
Function CheckTblEmpAllocations(FromTransDate As Date, TOtransDate As Date) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String
CheckTblEmpAllocations = False
sql = " SELECT RecordDate  "
sql = sql & " From dbo.TblEmpAllocations"
sql = sql & "  Where( RecordDate >=" & SQLDate(FromTransDate, True) & ") And (RecordDate <= " & SQLDate(TOtransDate, True) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckTblEmpAllocations = True
Else
CheckTblEmpAllocations = False
End If
End Function
Function CheckTblPaytAmortization(FromTransDate As Date, TOtransDate As Date) As Boolean
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim sql As String

CheckTblPaytAmortization = False
sql = " SELECT    RecorDate"
sql = sql & " From dbo.TblPaytAmortization"
sql = sql & " Where( RecorDate >=" & SQLDate(FromTransDate, True) & ") And (RecorDate <= " & SQLDate(TOtransDate, True) & ")"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
CheckTblPaytAmortization = True
Else
CheckTblPaytAmortization = False
End If
End Function

Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    sql = " SELECT * from  TblUsers "
sql = sql & " order by  UserName"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListUserAll.Clear
'    ListStoreSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
             
                ListUserAll.AddItem IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
               ListUserAll.ItemData(ListUserAll.NewIndex) = rs("UserID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

    'fil

    sql = " SELECT * from  TblGroupUsers "
 
 If SystemOptions.UserInterface = ArabicInterface Then
sql = sql & " order by  Name"
Else
sql = sql & " order by  Namee"
End If
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroup.Clear
'    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
             
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroup.AddItem IIf(IsNull(rs("Name").value), "", rs("Name").value)
            Else
                ListGroup.AddItem IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
            End If

            ListGroup.ItemData(ListGroup.NewIndex) = rs("ID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function
Function FillMylistData()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    sql = "SELECT     dbo.TblOpenClosPeriodDet2.ID, dbo.TblOpenClosPeriodDet2.Typ, dbo.TblOpenClosPeriodDet2.OPClsPerID, dbo.TblOpenClosPeriodDet2.UserID, "
    sql = sql & "                   dbo.TblUsers.UserName"
    sql = sql & "  FROM         dbo.TblOpenClosPeriodDet2 LEFT OUTER JOIN"
    sql = sql & "                   dbo.TblUsers ON dbo.TblOpenClosPeriodDet2.UserID = dbo.TblUsers.UserID"
    sql = sql & "   Where (dbo.TblOpenClosPeriodDet2.typ = 0) And (dbo.TblOpenClosPeriodDet2.OPClsPerID = " & val(TxtSerial1.Text) & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListUserAllSelected.Clear
'    ListStoreSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
             
                ListUserAllSelected.AddItem IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
               ListUserAllSelected.ItemData(ListUserAllSelected.NewIndex) = rs("UserID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

    'fil

    sql = " SELECT     dbo.TblOpenClosPeriodDet2.Typ, dbo.TblOpenClosPeriodDet2.OPClsPerID, dbo.TblOpenClosPeriodDet2.GroupID, dbo.TblGroupUsers.Name, "
    sql = sql & "                  dbo.TblGroupUsers.NameE , dbo.TblGroupUsers.ID"
    sql = sql & "   FROM         dbo.TblOpenClosPeriodDet2 LEFT OUTER JOIN"
    sql = sql & "                  dbo.TblGroupUsers ON dbo.TblOpenClosPeriodDet2.GroupID = dbo.TblGroupUsers.ID"
    sql = sql & "    WHERE     (dbo.TblOpenClosPeriodDet2.Typ = 1) AND (dbo.TblOpenClosPeriodDet2.OPClsPerID = " & val(TxtSerial1.Text) & ") "
 
 If SystemOptions.UserInterface = ArabicInterface Then
sql = sql & " order by  Name"
Else
sql = sql & " order by  Namee"
End If
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupSelected.Clear
'    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
             
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupSelected.AddItem IIf(IsNull(rs("Name").value), "", rs("Name").value)
            Else
                ListGroupSelected.AddItem IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
            End If

            ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = rs("GroupID").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function

Private Sub ChAll_Click()
If Me.ChAll.value = vbChecked Then
ChClose(0).value = vbChecked
ChClose(1).value = vbChecked
ChClose(2).value = vbChecked
ChClose(3).value = vbChecked
ChClose(4).value = vbChecked
ChClose(5).value = vbChecked
ChClose(6).value = vbChecked
ChClose(7).value = vbChecked
ElseIf Me.ChAll.value = vbUnchecked Then
'ChClose(0).value = vbUnchecked
ChClose(0).value = vbUnchecked
ChClose(1).value = vbUnchecked
ChClose(2).value = vbUnchecked
ChClose(3).value = vbUnchecked
ChClose(4).value = vbUnchecked
ChClose(5).value = vbUnchecked
ChClose(6).value = vbUnchecked
ChClose(7).value = vbUnchecked
End If
End Sub

Private Sub ChClose_Click(Index As Integer)
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
Select Case Index
Case 0
If Me.ChClose(0).value = vbChecked Then
Ch(0).value = xtpChecked
Ch(0).ForeColor = &HC00000
Else
Ch(0).value = xtpUnchecked
Ch(0).ForeColor = &HFF&
End If
Case 1
If Me.ChClose(1).value = vbChecked Then
Ch(1).value = xtpChecked
Ch(1).ForeColor = &HC00000
Else
Ch(1).value = xtpUnchecked
Ch(1).ForeColor = &HFF&
End If
Case 2
If Me.ChClose(2).value = vbChecked Then
Ch(2).ForeColor = &HC00000
Ch(2).value = xtpChecked
Else
Ch(2).ForeColor = &HFF&
Ch(2).value = xtpUnchecked
End If
Case 3
If Me.ChClose(3).value = vbChecked Then
Ch(3).value = xtpChecked
Ch(3).ForeColor = &HC00000
Else
Ch(3).value = xtpUnchecked
Ch(3).ForeColor = &HFF&
End If
Case 4
If Me.ChClose(4).value = vbChecked Then
Ch(4).value = xtpChecked
Ch(4).ForeColor = &HC00000
Else
Ch(4).value = xtpUnchecked
Ch(4).ForeColor = &HFF&
End If
Case 5
If Me.ChClose(5).value = vbChecked Then
Ch(5).value = xtpChecked
Ch(5).ForeColor = &HC00000
Else
Ch(5).value = xtpUnchecked
Ch(5).ForeColor = &HFF&
End If
Case 6
If Me.ChClose(6).value = vbChecked Then
Ch(6).value = xtpChecked
Ch(6).ForeColor = &HC00000
Else
Ch(6).value = xtpUnchecked
Ch(6).ForeColor = &HFF&
End If
Case 7
If Me.ChClose(7).value = vbChecked Then
Ch(7).value = xtpChecked
Ch(7).ForeColor = &HC00000
Else
Ch(7).value = xtpUnchecked
Ch(7).ForeColor = &HFF&
End If
End Select
End If
End Sub

Private Sub Check1_Click()
    Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.VSFlexGrid1
        
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid1

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With
         
    End If
End Sub
Sub Relain()
Dim i As Integer
Dim Yar As Integer
Dim mnth As Integer
Dim Fixed As Boolean
Dim Fixed2 As Boolean
Dim AlloCate As Boolean
Dim Amorti, Salar, Salar2 As Boolean
Dim Amorti2 As Boolean
Dim AlloCate2 As Boolean
Dim Approv, Approv2 As Boolean
Dim Bank, Bank2 As Boolean
Dim Purche, Purche2 As Boolean
Dim Sales, Sales2 As Boolean
  ChClose(0).value = vbUnchecked
    ChClose(1).value = vbUnchecked
    ChClose(2).value = vbUnchecked
    ChClose(3).value = vbUnchecked
    ChClose(4).value = vbUnchecked
    ChClose(5).value = vbUnchecked
    ChClose(6).value = vbUnchecked
    ChClose(7).value = vbUnchecked
    
Purche2 = True
Sales2 = True
Bank2 = True
Approv2 = True
AlloCate2 = True
Salar2 = True
Amorti2 = True
Amorti = False
Fixed2 = True
With Grid
For i = 1 To .Rows - 1
If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
If IsDate(.TextMatrix(i, .ColIndex("StartDate"))) And IsDate(.TextMatrix(i, .ColIndex("EndDate"))) Then
Yar = year(.TextMatrix(i, .ColIndex("StartDate")))
mnth = Month(.TextMatrix(i, .ColIndex("StartDate")))
Fixed = False

'''''''''''''''''''''''''//////////////
If CheckSales(.TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate"))) = True And Sales2 = True Then
Sales = True
Else
Sales = False
Sales2 = False
End If
'''''''''''''''''''''''''//////////////
If CheCkFixedAssest(.TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate"))) = True And Fixed2 = True Then
Fixed = True
Else
Fixed = False
Fixed2 = False
End If

'''''''''''''''''''''''''//////////////
If CheckPurche(.TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate"))) = True And Purche2 = True Then
Purche = True
Else
Purche = False
Purche2 = False
End If
'''''''''''''''''''''''''//////////////
If CheckLBankSettlement(.TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate"))) = True And Bank2 = True Then
Bank = True
Else
Bank = False
Bank2 = False
End If
'''''''''''''///////////////////////////
If CheCkSalary(Yar & mnth) = True And Salar2 = True Then
Salar = True
Else
Salar = False
Salar2 = False
End If
'''''//////////////
AlloCate = False
If CheckApprove(.TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate"))) = True And Approv2 = True Then
Approv = True
Else
Approv2 = False
Approv = False
End If
''/////////////////////
'''''//////////////
AlloCate = False
If CheckTblEmpAllocations(.TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate"))) = True And AlloCate2 = True Then
AlloCate = True
Else
AlloCate2 = False
AlloCate = False
End If
''/////////////////////
'''''''''''''///////////////////////////

If CheckTblPaytAmortization(.TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate"))) = True And Amorti2 = True Then
Amorti = True
Else
Amorti2 = False
Amorti = False
End If

End If
End If
Next i
End With

''''''''''''''''''''''''''''
If Sales = True Then
Ch(7).value = vbChecked
Ch(7).ForeColor = &HC00000
Else
Ch(7).value = vbUnchecked
Ch(7).ForeColor = &HFF&
End If
''''''''''''''''''''''''''''
If Purche = True Then
Ch(6).value = vbChecked
Ch(6).ForeColor = &HC00000
Else
Ch(6).value = vbUnchecked
Ch(6).ForeColor = &HFF&
End If
''''''''''''''''''''''''''''
If Bank = True Then
Ch(5).value = vbChecked
Ch(5).ForeColor = &HC00000
Else
Ch(5).value = vbUnchecked
Ch(5).ForeColor = &HFF&
End If
''''''''''''''''''''''''''''
If Approv = True Then
Ch(4).value = vbChecked
Ch(4).ForeColor = &HC00000
Else
Ch(4).value = vbUnchecked
Ch(4).ForeColor = &HFF&
End If
''''''''''''''''''''''''''''
If Salar = True Then
Ch(1).value = vbChecked
Ch(1).ForeColor = &HC00000
Else
Ch(1).value = vbUnchecked
Ch(1).ForeColor = &HFF&
End If
''''///////////////////////

''''''''''''''''''''''''''''
If Amorti = True Then
Ch(2).value = vbChecked
Ch(2).ForeColor = &HC00000
Else
Ch(2).value = vbUnchecked
Ch(2).ForeColor = &HFF&
End If
''''///////////////////////

If Fixed = True Then
Ch(0).value = vbChecked
Ch(0).ForeColor = &HC00000
Else
Ch(0).value = vbUnchecked
Ch(0).ForeColor = &HFF&
End If
''''///////////////////////
If AlloCate = True Then
Ch(3).value = vbChecked
Ch(3).ForeColor = &HC00000
Else
Ch(3).value = vbUnchecked
Ch(3).ForeColor = &HFF&
End If
End Sub
Private Sub ChkALL_Click()
    Dim i As Integer

    If ChkALL.value = vbChecked Then

        With Me.Grid
        
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = True
            Next i

        End With

    Else

        With Me.Grid

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("ch")) = False
            Next i

        End With
         
    End If
    
End Sub




Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub


Private Sub DcboEmpName2_Change()
DcboEmpName2_Click (0)
End Sub

Private Sub DcboEmpName2_Click(Area As Integer)
 If val(DcboEmpName2.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName2.BoundText, EmpCode
    TxtSearchCode2.Text = EmpCode
End Sub


Private Sub DcbType_Change()
DcbYear_Click (0)
End Sub

Private Sub DcbType_Click()
DcbType_Change
If val(DcbType.ListIndex) = 0 Then
Frame1.Visible = True
Frame2.Visible = False
ElseIf val(DcbType.ListIndex) = 1 Then
Frame2.Visible = True
Frame1.Visible = False
End If
End Sub

Private Sub DcbYear_Change()
DcbYear_Click (0)

End Sub

Private Sub DcbYear_Click(Area As Integer)

If Me.TxtModFlg.Text <> "R" Then
  Me.Grid.Clear flexClearScrollable, flexClearEverything
     Grid.Rows = 2
      Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
     VSFlexGrid1.Rows = 2
If val(DcbYear.BoundText) <> 0 Then
If val(DcbType.ListIndex) = 0 Then
filgrid val(DcbYear.BoundText)
ElseIf val(DcbType.ListIndex) = 1 Then
FilGrid12 val(DcbYear.BoundText)
End If
End If
End If
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
  
            
      FillMylist
      With DcbType
      If SystemOptions.UserInterface = ArabicInterface Then
      .Clear
      .AddItem "≈ﬁ›«· «·› —…"
      .AddItem "› Õ «·› —…"
      Else
        .Clear
      .AddItem "Locks Period"
      .AddItem "Opening  Period"
      End If
      End With
    conection = "select * from TblOpenClosPeriod order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCboUserName, My_SQL
    fill_combo DcboEmpName2, My_SQL
    fill_combo DcboEmpName, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.DcBranch
   ' Dcombos.GetEmployees Me.DcboEmpName
   ' Dcombos.GetEmployees Me.DcboEmpName2
     Dcombos.GetYarsData DcbYear
    BtnLast_Click
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
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
    Dim i As Integer
    Dim RsList As ADODB.Recordset
    If TxtModFlg = "E" Then
    StrSQL = "Delete From TblOpenClosPeriodDet1 Where OPClsPerID='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Delete From TblOpenClosPeriodDet2 Where OPClsPerID='" & val(TxtSerial1.Text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    End If

    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.DcBranch.BoundText)
    If val(DcbType.ListIndex) = 0 Then
    RsSavRec.Fields("EmpID").value = val(Me.DcboEmpName.BoundText)
    ElseIf val(DcbType.ListIndex) = 1 Then
    RsSavRec.Fields("EmpID").value = val(Me.DcboEmpName2.BoundText)
    End If
    RsSavRec.Fields("Yar").value = val(Me.DcbYear.BoundText)
    RsSavRec.Fields("TypAction").value = val(DcbType.ListIndex)
    If Ch(0).value = vbChecked Then
    RsSavRec.Fields("FlgAssets").value = 1
    Else
    RsSavRec.Fields("FlgAssets").value = 0
    End If
    
    If Me.ChAll.value = vbChecked Then
    RsSavRec.Fields("CloseAll").value = 1
    Else
    RsSavRec.Fields("CloseAll").value = 0
    End If
    
    If Ch(1).value = vbChecked Then
    RsSavRec.Fields("FlgSalary").value = 1
    Else
    RsSavRec.Fields("FlgSalary").value = 0
    End If
  
    If Ch(2).value = vbChecked Then
    RsSavRec.Fields("FlgAmortiz").value = 1
    Else
    RsSavRec.Fields("FlgAmortiz").value = 0
    End If
    If Ch(3).value = vbChecked Then
    RsSavRec.Fields("FlgAllocati").value = 1
    Else
    RsSavRec.Fields("FlgAllocati").value = 0
    End If
  
    If Ch(4).value = vbChecked Then
    RsSavRec.Fields("FlgApproved").value = 1
    Else
    RsSavRec.Fields("FlgApproved").value = 0
    End If
    
    If Ch(5).value = vbChecked Then
    RsSavRec.Fields("FlgBank").value = 1
    Else
    RsSavRec.Fields("FlgBank").value = 0
    End If
    
    If Ch(6).value = vbChecked Then
    RsSavRec.Fields("FlgPurchases").value = 1
    Else
    RsSavRec.Fields("FlgPurchases").value = 0
    End If
    If Ch(7).value = vbChecked Then
    RsSavRec.Fields("FlgSales").value = 1
    Else
    RsSavRec.Fields("FlgSales").value = 0
    End If
    If Ch(8).value = vbChecked Then
    RsSavRec.Fields("FlgUpdate").value = 1
    Else
    RsSavRec.Fields("FlgUpdate").value = 0
    End If
     If Ch(9).value = vbChecked Then
    RsSavRec.Fields("FlgPrint").value = 1
    Else
    RsSavRec.Fields("FlgPrint").value = 0
    End If
     If Ch(10).value = vbChecked Then
    RsSavRec.Fields("FlgDelete").value = 1
    Else
    RsSavRec.Fields("FlgDelete").value = 0
    End If
       If Ch(11).value = vbChecked Then
    RsSavRec.Fields("FlgAdded").value = 1
    Else
    RsSavRec.Fields("FlgAdded").value = 0
    End If
     If Opt(0).value = True Then
    RsSavRec.Fields("FlagAllUsers").value = 0
    ElseIf Opt(1).value = True Then
    RsSavRec.Fields("FlagAllUsers").value = 1
     ElseIf Opt(2).value = True Then
    RsSavRec.Fields("FlagAllUsers").value = 2
    End If
    
    ''///////////////
    If ChClose(0).value = vbChecked Then
    RsSavRec.Fields("CloseAssets").value = 1
    Else
    RsSavRec.Fields("CloseAssets").value = 0
    End If
    
     If ChClose(1).value = vbChecked Then
    RsSavRec.Fields("CloseSalary").value = 1
    Else
    RsSavRec.Fields("CloseSalary").value = 0
    End If
     If ChClose(2).value = vbChecked Then
    RsSavRec.Fields("CloseAmortiz").value = 1
    Else
    RsSavRec.Fields("CloseAmortiz").value = 0
    End If
     If ChClose(3).value = vbChecked Then
    RsSavRec.Fields("CloseAllocati").value = 1
    Else
    RsSavRec.Fields("CloseAllocati").value = 0
    End If
     If ChClose(4).value = vbChecked Then
    RsSavRec.Fields("CloseApproved").value = 1
    Else
    RsSavRec.Fields("CloseApproved").value = 0
    End If
     If ChClose(5).value = vbChecked Then
    RsSavRec.Fields("CloseBank").value = 1
    Else
    RsSavRec.Fields("CloseBank").value = 0
    End If
     If ChClose(6).value = vbChecked Then
    RsSavRec.Fields("ClosePurchases").value = 1
    Else
    RsSavRec.Fields("ClosePurchases").value = 0
    End If
     If ChClose(7).value = vbChecked Then
    RsSavRec.Fields("CloseSales").value = 1
    Else
    RsSavRec.Fields("CloseSales").value = 0
    End If
   
    
    '//////////
    
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblOpenClosPeriodDet1 Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim mnth As Integer
    Dim Yar As Integer
   If val(DcbType.ListIndex) = 0 Then
    With Grid
       For i = .FixedRows To .Rows - 1
     If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
       RsDevsub.AddNew
                RsDevsub("OPClsPerID").value = Me.TxtSerial1.Text
                RsDevsub("Typ").value = 0
                RsDevsub("Comment").value = IIf((.TextMatrix(i, .ColIndex("Comment"))) = "", Null, .TextMatrix(i, .ColIndex("Comment")))
                RsDevsub("StartDate").value = IIf((.TextMatrix(i, .ColIndex("StartDate"))) = "", Null, (.TextMatrix(i, .ColIndex("StartDate"))))
                RsDevsub("EndDate").value = IIf((.TextMatrix(i, .ColIndex("EndDate"))) = "", Null, .TextMatrix(i, .ColIndex("EndDate")))
      RsDevsub.update
      Yar = year(.TextMatrix(i, .ColIndex("StartDate")))
      mnth = Month(.TextMatrix(i, .ColIndex("StartDate")))
      UpdateFunctions .TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate")), Yar & mnth
      End If
     Next i
    End With
    ElseIf val(DcbType.ListIndex) = 1 Then 'open
        With VSFlexGrid1
       For i = .FixedRows To .Rows - 1
     If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked Then
       RsDevsub.AddNew
                RsDevsub("OPClsPerID").value = Me.TxtSerial1.Text
                RsDevsub("Typ").value = 1
                RsDevsub("Comment").value = IIf((.TextMatrix(i, .ColIndex("Comment"))) = "", Null, .TextMatrix(i, .ColIndex("Comment")))
                RsDevsub("StartDate").value = IIf((.TextMatrix(i, .ColIndex("StartDate"))) = "", Null, (.TextMatrix(i, .ColIndex("StartDate"))))
                RsDevsub("EndDate").value = IIf((.TextMatrix(i, .ColIndex("EndDate"))) = "", Null, .TextMatrix(i, .ColIndex("EndDate")))
      RsDevsub.update
          Yar = year(.TextMatrix(i, .ColIndex("StartDate")))
      mnth = Month(.TextMatrix(i, .ColIndex("StartDate")))
      UNUpdateFunctions .TextMatrix(i, .ColIndex("StartDate")), .TextMatrix(i, .ColIndex("EndDate")), Yar & mnth
      
      End If
     Next i
    End With
    End If
   '''''''''''''listUserID
  If ListUserAllSelected.ListCount > -1 Then
    Set RsList = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblOpenClosPeriodDet2 Where (1 = -1)"
    RsList.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With ListUserAllSelected
   For i = 0 To .ListCount - 1
                RsList.AddNew
                RsList("OPClsPerID").value = Me.TxtSerial1.Text
                RsList("UserID").value = .ItemData(i)
                RsList("Typ").value = 0
                RsList.update
            Next i
   End With
  RsList.Close
Set RsList = Nothing
        End If
 '''''''''''''listGroupID
   If ListGroupSelected.ListCount > -1 Then
    Set RsList = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblOpenClosPeriodDet2 Where (1 = -1)"
    RsList.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With ListGroupSelected
   For i = 0 To .ListCount - 1
                RsList.AddNew
                RsList("OPClsPerID").value = Me.TxtSerial1.Text
                RsList("GroupID").value = .ItemData(i)
                RsList("Typ").value = 1
                RsList.update
            Next i
   End With
  RsList.Close
Set RsList = Nothing
        End If
        
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
                Msg = Msg + "Do you want to enter another operation?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
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
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.Grid.Clear flexClearScrollable, flexClearEverything
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
  ' On Error GoTo ErrTrap
    Dim i As Integer
    ProgressBar1.Visible = True
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value): ProgressBar1.value = 20
    DcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), 0, RsSavRec.Fields("BranchID").value): ProgressBar1.value = 30
    Me.DcbYear.BoundText = IIf(IsNull(RsSavRec.Fields("Yar").value), "", RsSavRec.Fields("Yar").value): ProgressBar1.value = 40
    Me.DcbType.ListIndex = IIf(IsNull(RsSavRec.Fields("TypAction").value), -1, RsSavRec.Fields("TypAction").value): ProgressBar1.value = 50
    If val(DcbType.ListIndex) = 0 Then

    DcboEmpName.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value): ProgressBar1.value = 60
    Else
    DcboEmpName2.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    End If
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 70
If RsSavRec.Fields("FlgAssets").value = True Then
Ch(0).value = vbChecked: ProgressBar1.value = 80
Else
Ch(0).value = vbUnchecked: ProgressBar1.value = 80
End If

If RsSavRec.Fields("FlgSalary").value = True Then
Ch(1).value = vbChecked: ProgressBar1.value = 90
Else
Ch(1).value = vbUnchecked: ProgressBar1.value = 90
End If
If RsSavRec.Fields("FlgAmortiz").value = True Then
Ch(2).value = vbChecked: ProgressBar1.value = 100
Else
Ch(2).value = vbUnchecked: ProgressBar1.value = 100
End If
If RsSavRec.Fields("FlgAllocati").value = True Then
Ch(3).value = vbChecked: ProgressBar1.value = 10
Else
Ch(3).value = vbUnchecked: ProgressBar1.value = 10
End If
If RsSavRec.Fields("FlgApproved").value = True Then
Ch(4).value = vbChecked: ProgressBar1.value = 20
Else
Ch(4).value = vbUnchecked: ProgressBar1.value = 20
End If
If RsSavRec.Fields("FlgBank").value = True Then
Ch(5).value = vbChecked: ProgressBar1.value = 30
Else
Ch(5).value = vbUnchecked: ProgressBar1.value = 30
End If
If RsSavRec.Fields("FlgPurchases").value = True Then
Ch(6).value = vbChecked: ProgressBar1.value = 40
Else
Ch(6).value = vbUnchecked: ProgressBar1.value = 40
End If
If RsSavRec.Fields("FlgSales").value = True Then
Ch(7).value = vbChecked: ProgressBar1.value = 50
Else
Ch(7).value = vbUnchecked: ProgressBar1.value = 50
End If
If RsSavRec.Fields("FlgUpdate").value = True Then
Ch(8).value = vbChecked: ProgressBar1.value = 60
Else
Ch(8).value = vbUnchecked: ProgressBar1.value = 60
End If

If RsSavRec.Fields("FlgPrint").value = True Then
Ch(9).value = vbChecked: ProgressBar1.value = 70
Else
Ch(9).value = vbUnchecked: ProgressBar1.value = 70
End If
If RsSavRec.Fields("FlgDelete").value = True Then
Ch(10).value = vbChecked: ProgressBar1.value = 80
Else
Ch(10).value = vbUnchecked: ProgressBar1.value = 80
End If
If RsSavRec.Fields("FlgAdded").value = True Then
Ch(11).value = vbChecked: ProgressBar1.value = 90
Else
Ch(11).value = vbUnchecked: ProgressBar1.value = 90
End If
If RsSavRec.Fields("FlagAllUsers").value = 0 Then
Opt(0).value = True: ProgressBar1.value = 100
ElseIf RsSavRec.Fields("FlagAllUsers").value = 1 Then
Opt(1).value = True: ProgressBar1.value = 10
ElseIf RsSavRec.Fields("FlagAllUsers").value = 2 Then
Opt(2).value = True: ProgressBar1.value = 20
End If
''//////////
If RsSavRec.Fields("CloseAssets").value = True Then
ChClose(0).value = vbChecked: ProgressBar1.value = 30
Else
ChClose(0).value = vbUnchecked: ProgressBar1.value = 30
End If
If RsSavRec.Fields("CloseSalary").value = True Then
ChClose(1).value = vbChecked: ProgressBar1.value = 40
Else
ChClose(1).value = vbUnchecked: ProgressBar1.value = 40
End If
If RsSavRec.Fields("CloseAmortiz").value = True Then
ChClose(2).value = vbChecked: ProgressBar1.value = 50
Else
ChClose(2).value = vbUnchecked: ProgressBar1.value = 50
End If
If RsSavRec.Fields("CloseAllocati").value = True Then
ChClose(3).value = vbChecked: ProgressBar1.value = 60
Else
ChClose(3).value = vbUnchecked: ProgressBar1.value = 60
End If
If RsSavRec.Fields("CloseApproved").value = True Then
ChClose(4).value = vbChecked: ProgressBar1.value = 70
Else
ChClose(4).value = vbUnchecked: ProgressBar1.value = 70
End If
If RsSavRec.Fields("CloseBank").value = True Then
ChClose(5).value = vbChecked: ProgressBar1.value = 80
Else
ChClose(5).value = vbUnchecked: ProgressBar1.value = 80
End If
If RsSavRec.Fields("ClosePurchases").value = True Then
ChClose(6).value = vbChecked: ProgressBar1.value = 90
Else
ChClose(6).value = vbUnchecked: ProgressBar1.value = 90
End If
If RsSavRec.Fields("CloseSales").value = True Then
ChClose(7).value = vbChecked: ProgressBar1.value = 100
Else
ChClose(7).value = vbUnchecked: ProgressBar1.value = 100
End If

If RsSavRec.Fields("CloseAll").value = True Then
ChAll.value = vbChecked: ProgressBar1.value = 10
Else
ChAll.value = vbUnchecked: ProgressBar1.value = 10
End If
'
     LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 20
    LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 30
'     ' grid
    FullGrid
    FillMylistData
 ProgressBar1.Visible = False
 ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
ProgressBar1.value = 0
End Sub
  Sub FullGrid()
    Dim Rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
      Me.Grid.Clear flexClearScrollable, flexClearEverything
     Grid.Rows = 2
      Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
     VSFlexGrid1.Rows = 2
   Dim sql As String
   sql = "SELECT  * from  TblOpenClosPeriodDet1"
   sql = sql & " Where (OPClsPerID =" & val(TxtSerial1.Text) & " and Typ=0) "
   Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
       With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Comment")) = IIf(IsNull(Rs1("Comment").value), "", Rs1("Comment").value)
                   .TextMatrix(i, .ColIndex("StartDate")) = IIf(IsNull(Rs1("StartDate").value), "", Rs1("StartDate").value)
                   .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(Rs1("EndDate").value), "", Rs1("EndDate").value)
                    Rs1.MoveNext
             Next i
             .AutoSize 0, .Cols - 1, False
        End With
       Set rs2 = New ADODB.Recordset
   sql = "SELECT  * from  TblOpenClosPeriodDet1"
   sql = sql & " Where (OPClsPerID =" & val(TxtSerial1.Text) & " and Typ=1) "
   rs2.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
     rs2.MoveFirst
     End If
     
       With Me.VSFlexGrid1
                    For i = .FixedRows To rs2.RecordCount
                   .Rows = .FixedRows + rs2.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("Comment")) = IIf(IsNull(rs2("Comment").value), "", rs2("Comment").value)
                   .TextMatrix(i, .ColIndex("StartDate")) = IIf(IsNull(rs2("StartDate").value), "", rs2("StartDate").value)
                   .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(rs2("EndDate").value), "", rs2("EndDate").value)
                    rs2.MoveNext
             Next i
             .AutoSize 0, .Cols - 1, False
        End With
        
 End Sub



Function ChekGaridNotEmpty() As Boolean
Dim i As Integer
ChekGaridNotEmpty = False
With Grid
For i = 1 To .Rows - 1
 If .Cell(flexcpChecked, i, .ColIndex("ch")) = flexChecked And .TextMatrix(i, .ColIndex("StartDate")) <> "" Then
 ChekGaridNotEmpty = True
 Exit Function
 End If
 Next i
End With

End Function

Private Sub ISButton2_Click()
    If DcbYear.Text = "" Or val(Me.DcbYear.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«—  «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Year ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
          DcbYear.SetFocus
          Exit Sub
     End If
     If ChekGaridNotEmpty() = False Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
     Else
     MsgBox "Please Select Period"
     End If
     ChAll.Enabled = False
     Frame6.Enabled = False
     Exit Sub
     End If
      ChAll.Enabled = True
     Frame6.Enabled = True
Relain
End Sub

Private Sub Label10_Click()
If ListGroup.ListIndex = -1 Then Exit Sub
ListGroupSelected.AddItem ListGroup.List(ListGroup.ListIndex)
ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroup.ItemData(ListGroup.ListIndex)
End Sub

Private Sub Label3_Click()
If ListGroupSelected.ListIndex = -1 Then Exit Sub
ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
End Sub

Private Sub Label4_Click()
ListGroupSelected.Clear
End Sub

Private Sub Label5_Click()
If ListUserAllSelected.ListIndex = -1 Then Exit Sub
ListUserAllSelected.RemoveItem ListUserAllSelected.ListIndex
End Sub

Private Sub Label6_Click()
ListUserAllSelected.Clear
End Sub

Private Sub Label7_Click()
    Dim i As Integer
    ListUserAllSelected.Clear
    For i = 0 To Me.ListUserAll.ListCount - 1
  
        ListUserAllSelected.AddItem ListUserAll.List(i)
        ListUserAllSelected.ItemData(i) = ListUserAll.ItemData(i)
    Next i
End Sub

Private Sub Label8_Click()
If ListUserAll.ListIndex = -1 Then Exit Sub
ListUserAllSelected.AddItem ListUserAll.List(ListUserAll.ListIndex)
ListUserAllSelected.ItemData(ListUserAllSelected.NewIndex) = ListUserAll.ItemData(ListUserAll.ListIndex)
End Sub

Private Sub Label9_Click()
Dim i As Integer
ListGroupSelected.Clear
For i = 0 To Me.ListGroup.ListCount - 1
ListGroupSelected.AddItem ListGroup.List(i)
ListGroupSelected.ItemData(i) = ListGroup.ItemData(i)
Next i

End Sub



Private Sub Opt_Click(Index As Integer)
ListGroupSelected.Enabled = False
ListGroup.Enabled = False
Label10.Enabled = False
Label9.Enabled = False
Label4.Enabled = False
Label3.Enabled = False

ListUserAll.Enabled = False
ListUserAllSelected.Enabled = False
Label7.Enabled = False
Label8.Enabled = False
Label6.Enabled = False
Label5.Enabled = False
If Opt(0).value = True Then
ListUserAll.Enabled = True
ListUserAllSelected.Enabled = True
Label7.Enabled = True
Label8.Enabled = True
Label6.Enabled = True
Label5.Enabled = True
ElseIf Opt(1).value = True Then
ListGroupSelected.Enabled = True
ListGroup.Enabled = True
Label10.Enabled = True
Label9.Enabled = True
Label4.Enabled = True
Label3.Enabled = True
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub
  
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Sub UNUpdateFunctions(Optional StartDate As Date, Optional ToDate As Date, Optional sgn1 As String)
Dim sql As String
    sql = "update Notes  set LockedInterval=Null where NoteDate>=" & SQLDate(StartDate, True) & " and NoteDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
       
   sql = "update FixedAssetInstallments  set LockedInterval=Null where RecordDate>=" & SQLDate(StartDate, True) & " and RecordDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
      sql = "update ApprovalData  set LockedInterval=Null where SendTime>=" & SQLDate(StartDate, True) & " and SendTime<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
    sql = "update Notes  set LockedInterval=Null where NoteDate>=" & SQLDate(StartDate, True) & " and NoteDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
   sql = "update Notes1  set LockedInterval=Null where NoteDate>=" & SQLDate(StartDate, True) & " and NoteDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
   sql = "update notes_all  set LockedInterval=Null where NoteDate>=" & SQLDate(StartDate, True) & " and NoteDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
   sql = "update Transactions  set LockedInterval=Null where Transaction_Date>=" & SQLDate(StartDate, True) & " and Transaction_Date<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
   sql = "update TblEmpAllocations  set LockedInterval=Null where RecordDate>=" & SQLDate(StartDate, True) & " and RecordDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       
       DoEvents
   sql = "update TblPaytAmortization  set LockedInterval=Null where RecorDate>=" & SQLDate(StartDate, True) & " and RecorDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
    sql = "update emp_salary  set LockedInterval=Null where sgn='" & sgn1 & "'"
       Cn.Execute sql
       DoEvents
    sql = "update TblAccountIntervals  set OpenState=0 where StartDate=" & SQLDate(StartDate, True) & " and EndDate=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
     sql = "update TBLBankSettlement  set LockedInterval=0 where SettlementDT=" & SQLDate(StartDate, True) & " and SettlementDT=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
       DoEvents
End Sub
Sub UpdateFunctions(Optional StartDate As Date, Optional ToDate As Date, Optional sgn1 As String)
Dim sql As String
   sql = "update FixedAssetInstallments  set LockedInterval=1 where RecordDate>=" & SQLDate(StartDate, True) & " and RecordDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
      sql = "update ApprovalData  set LockedInterval=1 where SendTime>=" & SQLDate(StartDate, True) & " and SendTime<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
    sql = "update Notes  set LockedInterval=1 where NoteDate>=" & SQLDate(StartDate, True) & " and NoteDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
   sql = "update Notes1  set LockedInterval=1 where NoteDate>=" & SQLDate(StartDate, True) & " and NoteDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
   sql = "update notes_all  set LockedInterval=1 where NoteDate>=" & SQLDate(StartDate, True) & " and NoteDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
   sql = "update Transactions  set LockedInterval=1 where Transaction_Date>=" & SQLDate(StartDate, True) & " and Transaction_Date<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
   sql = "update TblEmpAllocations  set LockedInterval=1 where RecordDate>=" & SQLDate(StartDate, True) & " and RecordDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
   sql = "update TblPaytAmortization  set LockedInterval=1 where RecorDate>=" & SQLDate(StartDate, True) & " and RecorDate<=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
    sql = "update emp_salary  set LockedInterval=1 where sgn='" & sgn1 & "'"
       Cn.Execute sql
    sql = "update TblAccountIntervals  set OpenState=1 where StartDate=" & SQLDate(StartDate, True) & " and EndDate=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
   sql = "update TBLBankSettlement  set LockedInterval=1 where SettlementDT=" & SQLDate(StartDate, True) & " and SettlementDT=" & SQLDate(ToDate, True) & " "
       Cn.Execute sql
End Sub
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
  
    '---------------------- check if data Vaclete -----------------------
      If DcBranch.Text = "" Or val(Me.DcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
          DcBranch.SetFocus
            Exit Sub
     End If
          If DcbType.Text = "" Or val(Me.DcbType.ListIndex) = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·«Ã—«¡ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Kind of Action ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
          DcbType.SetFocus
          Exit Sub
     End If
         If DcbYear.Text = "" Or val(Me.DcbYear.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«—  «·”‰… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Year ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
          DcbYear.SetFocus
          Exit Sub
     End If
    If val(DcbType.ListIndex) = 0 Then
       If DcboEmpName.Text = "" Or val(Me.DcboEmpName.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«—  «·ﬁ«∆„ »«·«ﬁ›«· ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
          DcboEmpName.SetFocus
          Exit Sub
     End If
    ElseIf val(DcbType.ListIndex) = 1 Then
       If DcboEmpName2.Text = "" Or val(Me.DcboEmpName2.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«—  «·ﬁ«∆„ »«·› Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
          DcboEmpName2.SetFocus
          Exit Sub
     End If
    End If
    If val(DcbType.ListIndex) = 0 Then
    If Ch(0).value = vbChecked And Ch(1).value = vbChecked And Ch(2).value = vbChecked And Ch(3).value = vbChecked And Ch(4).value = vbChecked And Ch(5).value = vbChecked And Ch(6).value = vbChecked And Ch(7).value = vbChecked Then
    Else
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ «·« »⁄œ «ﬁ›«· ﬂ· «·Õ—ﬂ«  "
    Else
    MsgBox "Can Not Saved"
    End If
    Exit Sub
    End If
    End If
            '+++++++++++++++++++++++++++++++++++++++++++++++
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
'   StrVacName = IsRecExist("TblEmploymentModel", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtSerial1.text) & "'")
  ' If StrVacName <> "" Then
 '    Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·«”„ „‰ ﬁ»·"
  '     MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
  '    TxtVacName.SetFocus
 '     Exit Sub
'   End If
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text
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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
'  On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblOpenClosPeriod", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Private Sub TxtSearchCode2_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode2.Text, EmpID
        DcboEmpName2.BoundText = EmpID
    End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
' search for select id
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "ID=" & RecId, , adSearchForward, 1
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
 If Me.TxtModFlg.Text = "N" Then
 BtnLast_Click
' FiLLTXT
     
  Else
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
 FiLLTXT
     BtnLast_Click
 End If
End Sub
' delet sub
Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.Text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
            
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                RsSavRec.delete
                 StrSQL = "Delete From TblOpenClosPeriodDet1 Where OPClsPerID=" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                  StrSQL = "Delete From TblOpenClosPeriodDet2 Where OPClsPerID=" & val(TxtSerial1.Text) & ""
                   Cn.Execute StrSQL, , adExecuteNoRecords

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If
               
                Me.Grid.Clear flexClearScrollable, flexClearEverything
     Grid.Rows = 2
      Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
     VSFlexGrid1.Rows = 2
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
           Cn.Errors.Clear
    End Select

End Sub
' exit without save sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Me.TxtModFlg.Text <> "R" Then
        Select Case Me.TxtModFlg.Text
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
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)
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
                   RecId As String)
     FiLLRec
End Sub
Private Sub TxtModFlg_Change()
    If TxtModFlg.Text = "N" Then
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
       
        
        
    ElseIf TxtModFlg.Text = "R" Then
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtSerial1.Text <> "" Then
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
   ElseIf TxtModFlg.Text = "E" Then
   If Opt(0).value = True Then
   DcboEmpName.Enabled = False
   TxtSearchCode.Enabled = False
  ' DBCboClientName.Enabled = True
  ' TxtCustCode.Enabled = True
   Else
   DcboEmpName.Enabled = True
   TxtSearchCode.Enabled = True
 '  DBCboClientName.Enabled = False
'   TxtCustCode.Enabled = False
   End If
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
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
       
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveFirst
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnLast_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MoveLast

    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
        Grid.Rows = Grid.Rows + 1
        DcboEmpName2.BoundText = user_id
        Me.DCboUserName.BoundText = user_id
        DcboEmpName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.DcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
                    If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If
    End Select
End Sub
Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    Opt(0).value = True
      ChAll.Enabled = False
    ListUserAllSelected.Clear
    ListGroupSelected.Clear
    TxtModFlg.Text = "N"
    DcboEmpName2.BoundText = user_id
    Me.DCboUserName.BoundText = user_id
    DcboEmpName.BoundText = user_id
    Me.DcBranch.BoundText = branch_id
    DcbType.ListIndex = 0
    DcBranch.SetFocus
    Frame6.Enabled = False
    Ch(0).value = vbUnchecked
    Ch(1).value = vbUnchecked
    Ch(2).value = vbUnchecked
    Ch(3).value = vbUnchecked
    Ch(4).value = vbUnchecked
    Ch(5).value = vbUnchecked
    Ch(6).value = vbUnchecked
    Ch(7).value = vbUnchecked
    Ch(8).value = vbUnchecked
    Ch(9).value = vbUnchecked
    Ch(10).value = vbUnchecked
    Ch(11).value = vbUnchecked
    ChClose(0).value = vbUnchecked
    ChClose(1).value = vbUnchecked
    ChClose(2).value = vbUnchecked
    ChClose(3).value = vbUnchecked
    ChClose(4).value = vbUnchecked
    ChClose(5).value = vbUnchecked
    ChClose(6).value = vbUnchecked
    ChClose(7).value = vbUnchecked

     Me.Grid.Clear flexClearScrollable, flexClearEverything
     Grid.Rows = 2
      Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
     VSFlexGrid1.Rows = 2
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
      
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
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1.Text)
        Me.TxtModFlg.Text = "R"
    End If
    TxtModFlg = "R"
    If RsSavRec.RecordCount = 0 Then
        clear_all Me
       
        Exit Sub
    End If
BegnieWork:
    RsSavRec.MovePrevious
    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If
    
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
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
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
             .AddControl btnNew, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With
    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
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
        If Me.TxtModFlg.Text = "R" Then
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
Private Sub ISButton1_Click()
On Error GoTo ErrTrap
 '  If val(Me.TxtSerial1.text) <> 0 Then
 '      print_report
 '  End If
ErrTrap:
End Sub

Private Sub ChangeLang()
On Error GoTo ErrTrap
   
    Me.Caption = "Opening and Locks Periods  "
    Label1(2).Caption = Me.Caption
    lbl(4).Caption = "Trans ID"
    lbl(2).Caption = "Date"
    lbl(7).Caption = "Branch"
    lbl(5).Caption = "Year"
    Frame1.Caption = "In Case of Locks"
    Frame2.Caption = "In Case of Opening"
    Frame7.Caption = "Open Options"
    lbl(1).Caption = "Open periods"
    lbl(3).Caption = "Locks periods"
    ISButton2.Caption = "Run Auto"
    lbl(0).Caption = "Employee"
    lbl(10).Caption = "Employee"
    Ch(0).RightToLeft = False
     ChAll.Caption = "Close All"
     Me.ChAll.RightToLeft = False
Ch(1).RightToLeft = False
Ch(2).RightToLeft = False
Ch(3).RightToLeft = False
Ch(4).RightToLeft = False
Ch(5).RightToLeft = False
Ch(6).RightToLeft = False
Ch(7).RightToLeft = False
Ch(8).RightToLeft = False
Ch(9).RightToLeft = False
Ch(10).RightToLeft = False
Ch(11).RightToLeft = False
ChkALL.RightToLeft = False
Check1.RightToLeft = False
Check1.Caption = "Select All"
ChkALL.Caption = "Select All"
Opt(0).RightToLeft = False
Opt(1).RightToLeft = False
Opt(2).RightToLeft = False
Opt(2).Caption = "All Users"
Opt(1).Caption = "Group"
Opt(0).Caption = "User"
Ch(8).Caption = "Can Edited"
Ch(9).Caption = "Can Print"
Ch(10).Caption = "Can Delete"
Ch(11).Caption = "Can Added"
Ch(7).Caption = "Sales not Received"
Ch(6).Caption = "Purchases not Received"
Ch(5).Caption = "Bank settlements"
Ch(4).Caption = "Doc.To.Approve"
Ch(1).Caption = "Entitlement Salary"
Ch(0).Caption = "Installments Assets"
Ch(2).Caption = "Payment Amortization"
Ch(3).Caption = "Registration Allocations"
lbl(12).Caption = "Type Action"
lbl(6).Caption = "Implementation manual locks after Auto locks"
    ''''''''''''''''''''''''''''''''''''''' next
    ChClose(0).RightToLeft = False
    ChClose(1).RightToLeft = False
    ChClose(2).RightToLeft = False
    ChClose(3).RightToLeft = False
    ChClose(4).RightToLeft = False
    ChClose(5).RightToLeft = False
    ChClose(6).RightToLeft = False
    ChClose(7).RightToLeft = False
    ChClose(0).Caption = "Manual Locks"
    ChClose(1).Caption = "Manual Locks"
    ChClose(2).Caption = "Manual Locks"
    ChClose(3).Caption = "Manual Locks"
    ChClose(4).Caption = "Manual Locks"
    ChClose(5).Caption = "Manual Locks"
    ChClose(6).Caption = "Manual Locks"
    ChClose(7).Caption = "Manual Locks"
   ''///////////
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "NO. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
   ' ISButton6.Caption = "Delet Select"
  '  ISButton4.Caption = "Delet All"
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
   .TextMatrix(0, .ColIndex("ch")) = "Select"
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("StartDate")) = "From"
        .TextMatrix(0, .ColIndex("EndDate")) = "TO"
        .TextMatrix(0, .ColIndex("Comment")) = "Remarks"
    End With
       With Me.VSFlexGrid1
   .TextMatrix(0, .ColIndex("ch")) = "Select"
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("StartDate")) = "From"
        .TextMatrix(0, .ColIndex("EndDate")) = "TO"
        .TextMatrix(0, .ColIndex("Comment")) = "Remarks"
    End With
ErrTrap:
End Sub

Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
'  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblOpenClosPeriod"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ end






