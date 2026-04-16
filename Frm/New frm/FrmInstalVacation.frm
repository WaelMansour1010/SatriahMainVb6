VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmInstalVacation 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmInstalVacation.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
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
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   4395
      Left            =   0
      TabIndex        =   40
      Top             =   2880
      Width           =   14235
      _cx             =   25109
      _cy             =   7752
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmInstalVacation.frx":6852
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
      ExplorerBar     =   3
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
         TabIndex        =   41
         Top             =   960
         Visible         =   0   'False
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      TabIndex        =   26
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmInstalVacation.frx":6A0D
      Left            =   15480
      List            =   "FrmInstalVacation.frx":6A1D
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   14505
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
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
         ButtonImage     =   "FrmInstalVacation.frx":6A36
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
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
         ButtonImage     =   "FrmInstalVacation.frx":6DD0
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
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
         ButtonImage     =   "FrmInstalVacation.frx":716A
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   23
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
         ButtonImage     =   "FrmInstalVacation.frx":7504
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«—’œÂ ≈›  «ÕÌ"
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
         TabIndex        =   24
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmInstalVacation.frx":789E
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4215
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   14235
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   1335
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   14055
         Begin VB.TextBox TxtAbcence 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6240
            TabIndex        =   66
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox TXtVacBalance 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   11760
            TabIndex        =   63
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtVacWithoutSal 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   9000
            TabIndex        =   65
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   11760
            TabIndex        =   54
            Top             =   240
            Width           =   975
         End
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   255
            Index           =   0
            Left            =   12720
            TabIndex        =   47
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "·„ÊŸ› „Õœœ"
            BackColor       =   16761024
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Bindings        =   "FrmInstalVacation.frx":8CA3
            Height          =   315
            Left            =   9000
            TabIndex        =   48
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
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   255
            Index           =   3
            Left            =   12720
            TabIndex        =   49
            Top             =   600
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "·ﬂ· «·„ÊŸ›Ì‰"
            BackColor       =   16761024
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   255
            Index           =   1
            Left            =   7440
            TabIndex        =   50
            Top             =   240
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "·›—⁄ „Õœœ"
            BackColor       =   16761024
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSelBranch 
            Bindings        =   "FrmInstalVacation.frx":8CB8
            Height          =   315
            Left            =   4560
            TabIndex        =   51
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
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   255
            Index           =   2
            Left            =   3240
            TabIndex        =   52
            Top             =   240
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "·≈œ«—… „ÕœœÂ"
            BackColor       =   16761024
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbDepartment 
            Bindings        =   "FrmInstalVacation.frx":8CCD
            Height          =   315
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3015
            _ExtentX        =   5318
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
         Begin Dynamic_Byte.NourHijriCal BeginDateH 
            Height          =   255
            Left            =   4560
            TabIndex        =   55
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
         End
         Begin MSComCtl2.DTPicker BeginDate 
            Height          =   315
            Left            =   6240
            TabIndex        =   56
            Top             =   600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal LastDateH 
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   450
         End
         Begin MSComCtl2.DTPicker LastDate 
            Height          =   315
            Left            =   1680
            TabIndex        =   59
            Top             =   600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   315
            Left            =   120
            TabIndex        =   67
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   960
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   556
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmInstalVacation.frx":8CE2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "€Ì«»"
            Height          =   285
            Index           =   6
            Left            =   7800
            TabIndex        =   64
            Top             =   960
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã«“… »œÊ‰ —« »"
            Height          =   285
            Index           =   5
            Left            =   10440
            TabIndex        =   62
            Top             =   960
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ ≈Ã«“…"
            Height          =   285
            Index           =   3
            Left            =   12720
            TabIndex        =   61
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ— „»«‘—…"
            Height          =   285
            Index           =   1
            Left            =   3240
            TabIndex        =   60
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ê· „»«‘—…"
            Height          =   285
            Index           =   0
            Left            =   7680
            TabIndex        =   57
            Top             =   600
            Width           =   1245
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   14055
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11400
            TabIndex        =   12
            Top             =   240
            Width           =   1455
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
            Height          =   255
            Left            =   6840
            TabIndex        =   13
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmInstalVacation.frx":F544
            Height          =   315
            Left            =   120
            TabIndex        =   1
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
            TabIndex        =   14
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   100597761
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   285
            Index           =   4
            Left            =   12960
            TabIndex        =   17
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   255
            Left            =   5280
            TabIndex        =   16
            Top             =   240
            Width           =   615
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
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   27
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
      TabIndex        =   28
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
      Height          =   2145
      Left            =   0
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6960
      Width           =   14235
      _cx             =   25109
      _cy             =   3784
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
         TabIndex        =   31
         Top             =   480
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   1680
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   30
         Top             =   1080
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   2
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
            ButtonImage     =   "FrmInstalVacation.frx":F559
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   4
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
            ButtonImage     =   "FrmInstalVacation.frx":15DBB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   3
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
            ButtonImage     =   "FrmInstalVacation.frx":16155
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   5
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
            ButtonImage     =   "FrmInstalVacation.frx":1C9B7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   6
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
            ButtonImage     =   "FrmInstalVacation.frx":1CD51
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   720
            TabIndex        =   7
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
            ButtonImage     =   "FrmInstalVacation.frx":1D2EB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   4080
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmInstalVacation.frx":1D685
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   2400
            TabIndex        =   69
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
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
            ButtonImage     =   "FrmInstalVacation.frx":23EE7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9600
         TabIndex        =   36
         Top             =   600
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
         Left            =   7560
         TabIndex        =   37
         ToolTipText     =   "Õ–› «·’› «·Õ«·Ì"
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·’› «·Õ«·Ì"
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
         ButtonImage     =   "FrmInstalVacation.frx":24281
         ButtonImageDisabled=   "FrmInstalVacation.frx":2AAE3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton ISButton4 
         Height          =   330
         Left            =   5880
         TabIndex        =   38
         ToolTipText     =   "Õ–› «·ﬂ·"
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› «·ﬂ· "
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
         ButtonImage     =   "FrmInstalVacation.frx":49CCD
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
         Width           =   13965
         _cx             =   24633
         _cy             =   1005
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   6
         ChildSpacing    =   4
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
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   12840
         TabIndex        =   39
         Top             =   600
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
            Picture         =   "FrmInstalVacation.frx":5052F
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInstalVacation.frx":508C9
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInstalVacation.frx":50C63
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInstalVacation.frx":50FFD
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInstalVacation.frx":51397
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInstalVacation.frx":51731
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInstalVacation.frx":51ACB
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmInstalVacation.frx":52065
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   42
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
      ButtonImage     =   "FrmInstalVacation.frx":523FF
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   45
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
      ButtonImage     =   "FrmInstalVacation.frx":58C61
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   46
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
      ButtonImage     =   "FrmInstalVacation.frx":5F4C3
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
      TabIndex        =   43
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmInstalVacation"
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
 Dim ii As Long
 Public LngRow As Double
Public LngCol As Double

Sub SaveInformationVacation(Optional TypeVacation As Integer = 0, Optional EmpID As Integer = 0, Optional NoDay As Double = 0)
Dim sql As String
Dim str As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
str = "«—’œÂ ≈›  «ÕÌ…"
Else
str = "Balances Opening"
End If
sql = "select * from TblInforVacatiom where (1=-1)"
    Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Rs7.AddNew
      Rs7("InstVacaID").value = val(TxtSerial1.text)
      Rs7("EmpID").value = EmpID
      Rs7("NoDay").value = (NoDay)
      Rs7("RecordDate").value = XPDtbTrans.value
      Rs7("RecordDateH").value = Txt_DateHigri.value
      Rs7("TypeVacation").value = TypeVacation
      Rs7("Remarks").value = str
      Rs7.update
End Sub
Sub SaveVacation(Optional EmpID As Integer = 0, Optional NoDay As Double = 0)
Dim sql As String
Dim str As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
sql = "select * from tblVacationData where (1=-1)"
    Rs7.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Rs7.AddNew
      Rs7("InstVacaID").value = val(TxtSerial1.text)
      Rs7("EmpID").value = EmpID
      Rs7("Value").value = (NoDay)
      Rs7("ExpectedacationDate").value = XPDtbTrans.value
      Rs7("ExpectedacationDateH").value = Txt_DateHigri.value
      If SystemOptions.UserInterface = ArabicInterface Then
str = "«—’œÂ ≈›  «ÕÌ…"
Else
str = "Balances Opening"
End If
Rs7("Remark").value = str
      Rs7.update
End Sub



Private Sub BeginDate_Change()
  If Me.TxtModFlg.text <> "R" Then
              BeginDateH.value = ToHijriDate(BeginDate.value)
              LastDate.value = BeginDate.value
   End If
End Sub



Private Sub BeginDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
  VBA.Calendar = vbCalGreg
            BeginDate.value = ToGregorianDate(BeginDateH.value)
            LastDateH.value = BeginDateH.value
 End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
    Dim sql As String
Dim Rs5 As ADODB.Recordset
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
Set Rs5 = New ADODB.Recordset
sql = "SELECT     Contract_date, DateH, Emp_id"
sql = sql & " From dbo.Contract"
sql = sql & " WHERE     (Emp_id = " & val(DcboEmpName.BoundText) & ")"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
BeginDate.value = IIf(IsNull(Rs5("Contract_date").value), BeginDate.value, Rs5("Contract_date").value)
BeginDateH.value = IIf(IsNull(Rs5("DateH").value), BeginDateH.value, Rs5("DateH").value)
End If
End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 45
      FrmEmployeeSearch.show
    End If
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
    conection = "select * from TblInstalVacation order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me

    
    
    If SystemOptions.OpeningEmployeeShowAll = True Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                                    My_SQL = "SELECT Emp_ID,Emp_Name From TblEmployee    "
                                    My_SQL = My_SQL + " Order By Emp_Name ASC"
                        Else
                                    My_SQL = "SELECT Emp_ID,Emp_Namee From TblEmployee    "
                                    My_SQL = My_SQL + " Order By Emp_Namee ASC"
                        End If
    Else
    
                      If SystemOptions.UserInterface = ArabicInterface Then
                                 My_SQL = "SELECT Emp_ID,Emp_Name From TblEmployee WHERE     (BignDateWork IS NULL)   "
                                 My_SQL = My_SQL + " Order By Emp_Name ASC"
                       Else
                                 My_SQL = "SELECT Emp_ID,Emp_Namee From TblEmployee WHERE     (BignDateWork IS NULL)   "
                                 My_SQL = My_SQL + " Order By Emp_Namee ASC"
                       End If
                    
    
    End If
    
    fill_combo DcboEmpName, My_SQL
    
   'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCboUserName, My_SQL
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetBranches Me.DcbSelBranch
    'Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmpDepartments Me.DcbDepartment
    BtnLast_Click
    DcboEmpName.Enabled = False
TxtSearchCode.Enabled = False
DcbSelBranch.Enabled = False
DcbDepartment.Enabled = False
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
    Dim i As Integer
    Dim sql As String
    If TxtModFlg = "E" Then
    With Me.Grid
    For i = 1 To .rows - 1
    If CheckEmp(val(.TextMatrix(i, .ColIndex("EmpID")))) = False Then
    StrSQL = "Delete From TblInstalVacationDet Where InslVaID=" & val(TxtSerial1.text) & " and EmpID=" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
    Cn.Execute StrSQL, , adExecuteNoRecords
   sql = "update TblEmployee set   BignDateWork =null ,IssueDateH=null  where Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
   Cn.Execute sql
   sql = "update TblEmployee set   lastHolidaydate = null ,lastHolidaydateH=null where Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
   Cn.Execute sql
    End If
    Next i
    End With
      StrSQL = "Delete From TblInforVacatiom Where InstVacaID=" & val(TxtSerial1.text) & ""
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From tblVacationData Where InstVacaID=" & val(TxtSerial1.text) & ""
    Cn.Execute StrSQL, , adExecuteNoRecords
    End If

    RsSavRec.Fields("RecordM").value = XPDtbTrans.value
    RsSavRec.Fields("RecordH").value = Me.Txt_DateHigri.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    RsSavRec.Fields("BranchSelectID").value = val(Me.DcbSelBranch.BoundText)
    RsSavRec.Fields("EmpID").value = val(Me.DcboEmpName.BoundText)
    RsSavRec.Fields("DeptID").value = val(Me.DcbDepartment.BoundText)
    RsSavRec.Fields("BeginDate").value = BeginDate.value
    RsSavRec.Fields("BeginDateH").value = Me.BeginDateH.value
    RsSavRec.Fields("LastDate").value = LastDate.value
    RsSavRec.Fields("LastDateH").value = Me.LastDateH.value
    RsSavRec.Fields("VacBalance").value = val(Me.TXtVacBalance.text)
    RsSavRec.Fields("VacWithoutSal").value = val(Me.TxtVacWithoutSal.text)
    RsSavRec.Fields("Abcence").value = val(Me.TxtAbcence.text)
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    If opt(0).value = True Then
    RsSavRec.Fields("TypeSelect").value = 0
    ElseIf opt(1).value = True Then
    RsSavRec.Fields("TypeSelect").value = 1
       ElseIf opt(2).value = True Then
    RsSavRec.Fields("TypeSelect").value = 2
       ElseIf opt(3).value = True Then
    RsSavRec.Fields("TypeSelect").value = 3
    Else
    RsSavRec.Fields("TypeSelect").value = Null
    End If
    ''/////
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    ' save grid
    Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblInstalVacationDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Grid
       For i = .FixedRows To .rows - 1
     If val(.TextMatrix(i, .ColIndex("EmpID"))) <> 0 Then
     If CheckEmp(val(.TextMatrix(i, .ColIndex("EmpID")))) = False Then
                RsDevsub.AddNew
                RsDevsub("InslVaID").value = Me.TxtSerial1.text
                RsDevsub("EmpID").value = IIf((.TextMatrix(i, .ColIndex("EmpID"))) = "", Null, .TextMatrix(i, .ColIndex("EmpID")))
                RsDevsub("BeginDate").value = IIf((.TextMatrix(i, .ColIndex("BeginDate"))) = "", Null, .TextMatrix(i, .ColIndex("BeginDate")))
                RsDevsub("LastDate").value = IIf((.TextMatrix(i, .ColIndex("LastDate"))) = "", Null, .TextMatrix(i, .ColIndex("LastDate")))
                RsDevsub("VacBalance").value = IIf((.TextMatrix(i, .ColIndex("VacBalance"))) = "", Null, .TextMatrix(i, .ColIndex("VacBalance")))
                RsDevsub("VacWithoutSal").value = IIf((.TextMatrix(i, .ColIndex("VacWithoutSal"))) = "", Null, .TextMatrix(i, .ColIndex("VacWithoutSal")))
                RsDevsub("Abcence").value = IIf((.TextMatrix(i, .ColIndex("Abcence"))) = "", Null, .TextMatrix(i, .ColIndex("Abcence")))
                RsDevsub("BeginDateH").value = IIf((.TextMatrix(i, .ColIndex("BeginDateH"))) = "", Null, .TextMatrix(i, .ColIndex("BeginDateH")))
                RsDevsub("LastDateH").value = IIf((.TextMatrix(i, .ColIndex("LastDateH"))) = "", Null, .TextMatrix(i, .ColIndex("LastDateH")))
                
                If val(((.TextMatrix(i, .ColIndex("VacWithoutSal"))))) <> 0 Then
                SaveInformationVacation 0, val(.TextMatrix(i, .ColIndex("EmpID"))), val(((.TextMatrix(i, .ColIndex("VacWithoutSal")))))
                End If
                If val(((.TextMatrix(i, .ColIndex("Abcence"))))) <> 0 Then
                SaveInformationVacation 1, val(.TextMatrix(i, .ColIndex("EmpID"))), val(((.TextMatrix(i, .ColIndex("Abcence")))))
                End If
                 If val(((.TextMatrix(i, .ColIndex("VacBalance"))))) <> 0 Then
                SaveVacation val(.TextMatrix(i, .ColIndex("EmpID"))), val(((.TextMatrix(i, .ColIndex("VacBalance")))))
                End If
                 sql = "update TblEmployee set   BignDateWork =  " & SQLDate((.TextMatrix(i, .ColIndex("BeginDate"))), True) & " ,IssueDateH='" & .TextMatrix(i, .ColIndex("BeginDateH")) & "'  where Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
                Cn.Execute sql
                sql = "update TblEmployee set   lastHolidaydate =  " & SQLDate((.TextMatrix(i, .ColIndex("LastDate"))), True) & " ,lastHolidaydateH='" & .TextMatrix(i, .ColIndex("LastDateH")) & "' where Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
                Cn.Execute sql
                
                sql = "update TblEmployee set   balanceH3 =  " & val(.TextMatrix(i, .ColIndex("VacBalance"))) & "  where Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
                Cn.Execute sql

                
      RsDevsub.update
      End If
      End If
     Next i
    End With
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " Saved... " & CHR(13)
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
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value): ProgressBar1.value = 10
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordM").value), Date, RsSavRec.Fields("RecordM").value): ProgressBar1.value = 20
    Txt_DateHigri.value = IIf(IsNull(RsSavRec.Fields("RecordH").value), "", RsSavRec.Fields("RecordH").value): ProgressBar1.value = 30
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value): ProgressBar1.value = 40
    DcboEmpName.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value): ProgressBar1.value = 50
    DcbSelBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchSelectID").value), "", RsSavRec.Fields("BranchSelectID").value): ProgressBar1.value = 60
    DcbDepartment.BoundText = IIf(IsNull(RsSavRec.Fields("DeptID").value), "", RsSavRec.Fields("DeptID").value): ProgressBar1.value = 70
If (IsNull(RsSavRec.Fields("TypeSelect").value)) Then
ElseIf val(RsSavRec.Fields("TypeSelect").value) = 0 Then
opt(0).value = True: ProgressBar1.value = 80
ElseIf val(RsSavRec.Fields("TypeSelect").value) = 1 Then
opt(1).value = True: ProgressBar1.value = 90
ElseIf val(RsSavRec.Fields("TypeSelect").value) = 2 Then
opt(2).value = True: ProgressBar1.value = 100
ElseIf val(RsSavRec.Fields("TypeSelect").value) = 3 Then
opt(3).value = True: ProgressBar1.value = 10
End If


     ''''''''''''''''
     DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value): ProgressBar1.value = 20
     Me.TXtVacBalance.text = IIf(IsNull(RsSavRec.Fields("VacBalance").value), "", RsSavRec.Fields("VacBalance").value): ProgressBar1.value = 30
     Me.TxtVacWithoutSal.text = IIf(IsNull(RsSavRec.Fields("VacWithoutSal").value), "", RsSavRec.Fields("VacWithoutSal").value): ProgressBar1.value = 40
     Me.TxtAbcence.text = IIf(IsNull(RsSavRec.Fields("Abcence").value), "", RsSavRec.Fields("Abcence").value): ProgressBar1.value = 50
     BeginDate.value = IIf(IsNull(RsSavRec.Fields("BeginDate").value), Date, RsSavRec.Fields("BeginDate").value): ProgressBar1.value = 60
     BeginDateH.value = IIf(IsNull(RsSavRec.Fields("BeginDateH").value), "", RsSavRec.Fields("BeginDateH").value): ProgressBar1.value = 70
     LastDate.value = IIf(IsNull(RsSavRec.Fields("LastDate").value), Date, RsSavRec.Fields("LastDate").value): ProgressBar1.value = 80
     LastDateH.value = IIf(IsNull(RsSavRec.Fields("LastDateH").value), "", RsSavRec.Fields("LastDateH").value): ProgressBar1.value = 90
    
     LabCurrRec.Caption = RsSavRec.AbsolutePosition: ProgressBar1.value = 100
     LabCountRec.Caption = RsSavRec.RecordCount: ProgressBar1.value = 10
     ' grid
    FullGrid
 ProgressBar1.Visible = False
 ProgressBar1.value = 0
ErrTrap:
  ProgressBar1.Visible = False
 ProgressBar1.value = 0
End Sub
  Sub FullGrid()
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset
   Dim sql As String
   sql = "SELECT     dbo.TblInstalVacationDet.ID, dbo.TblInstalVacationDet.InslVaID, dbo.TblInstalVacationDet.BeginDate, dbo.TblInstalVacationDet.LastDate, "
   sql = sql & "                   dbo.TblInstalVacationDet.VacBalance, dbo.TblInstalVacationDet.VacWithoutSal, dbo.TblInstalVacationDet.Abcence, dbo.TblInstalVacationDet.BeginDateH,"
   sql = sql & "                   dbo.TblInstalVacationDet.LastDateH , dbo.TblInstalVacationDet.empid, dbo.TblEmployee.emp_name, dbo.TblEmployee.fullcode, dbo.TblEmployee.Emp_Namee"
   sql = sql & "  FROM         dbo.TblInstalVacationDet LEFT OUTER JOIN"
   sql = sql & "                   dbo.TblEmployee ON dbo.TblInstalVacationDet.EmpID = dbo.TblEmployee.Emp_ID"
   sql = sql & "  Where (dbo.TblInstalVacationDet.InslVaID =" & val(TxtSerial1.text) & ")"

   Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
       With Me.Grid
                    For i = .FixedRows To Rs1.RecordCount
                   .rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(Rs1("EmpID").value), "", Rs1("EmpID").value)
                   .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                   .TextMatrix(i, .ColIndex("BeginDate")) = IIf(IsNull(Rs1("BeginDate").value), "", Rs1("BeginDate").value)
                   .TextMatrix(i, .ColIndex("LastDate")) = IIf(IsNull(Rs1("LastDate").value), "", Rs1("LastDate").value)
                   .TextMatrix(i, .ColIndex("BeginDateH")) = IIf(IsNull(Rs1("BeginDateH").value), "", Rs1("BeginDateH").value)
                   .TextMatrix(i, .ColIndex("LastDateH")) = IIf(IsNull(Rs1("LastDateH").value), "", Rs1("LastDateH").value)
                   
                      .TextMatrix(i, .ColIndex("VacBalance")) = IIf(IsNull(Rs1("VacBalance").value), "", Rs1("VacBalance").value)
                   .TextMatrix(i, .ColIndex("VacWithoutSal")) = IIf(IsNull(Rs1("VacWithoutSal").value), "", Rs1("VacWithoutSal").value)
                   .TextMatrix(i, .ColIndex("Abcence")) = IIf(IsNull(Rs1("Abcence").value), "", Rs1("Abcence").value)
                   
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
                   End If
                    Rs1.MoveNext
             Next i
        End With
     
        Exit Sub
 End Sub

Private Sub Grid_CellButtonClick(ByVal row As Long, ByVal Col As Long)
If Me.TxtModFlg.text <> "R" Then
With Grid
Select Case .ColKey(Col)


        Case "BeginDate"
        LngRow = row
        LngCol = Col
       
        Load FrmDateOpProject
        FrmDateOpProject.index = 1
        FrmDateOpProject.show vbModal
        
          Case "BeginDateH"
        LngRow = row
        LngCol = Col
        Load FrmDateOpProject
        FrmDateOpProject.index = 1
        FrmDateOpProject.show vbModal
        
          Case "LastDate"
        LngRow = row
        LngCol = Col
        Load FrmDateOpProject
        FrmDateOpProject.index = 1
        FrmDateOpProject.show vbModal
        
          Case "LastDateH"
        LngRow = row
        LngCol = Col
        Load FrmDateOpProject
        FrmDateOpProject.index = 1
        FrmDateOpProject.show vbModal
       End Select
      End With
    End If
End Sub

Private Sub Grid_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
 With Me.Grid

        Select Case .ColKey(Col)
                               Case "BeginDate"
.ColComboList(.ColIndex("BeginDate")) = "..."
                              Case "BeginDateH"
.ColComboList(.ColIndex("BeginDateH")) = "..."
                              Case "LastDate"
.ColComboList(.ColIndex("LastDate")) = "..."
                              Case "LastDateH"
.ColComboList(.ColIndex("LastDateH")) = "..."
End Select
End With
End Sub

Private Sub ISButton2_Click()

If opt(0).value = True Then
If val(DcboEmpName.BoundText) <> 0 Then
filgrid1
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·„ÊŸ›"
Else
MsgBox "Please Select Employee"
End If
DcboEmpName.SetFocus
Exit Sub
End If
ElseIf opt(1).value = True Then
If val(DcbSelBranch.BoundText) <> 0 Then
filgrid2 1
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·›—⁄"
Else
MsgBox "Please Select Branch"
End If
DcbSelBranch.SetFocus
Exit Sub
End If

ElseIf opt(2).value = True Then
If val(DcbDepartment.BoundText) <> 0 Then
filgrid2 2
Else
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ ≈Œ Ì«— «·«œ«—…"
Else
MsgBox "Please Select Management"
End If
DcbDepartment.SetFocus
Exit Sub
End If

ElseIf opt(3).value = True Then
filgrid2 3
End If

End Sub
Sub filgrid1()
Dim k As Integer
Dim i As Integer

With Grid
k = .rows - 1
.rows = .rows + 1
Do While k < (.rows - 1)
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("EmpID")) = DcboEmpName.BoundText
.TextMatrix(k, .ColIndex("Fullcode")) = TxtSearchCode.text
.TextMatrix(k, .ColIndex("name")) = DcboEmpName.text
.TextMatrix(k, .ColIndex("BeginDate")) = BeginDate.value
.TextMatrix(k, .ColIndex("BeginDateH")) = BeginDateH.value
.TextMatrix(k, .ColIndex("LastDate")) = LastDate.value
.TextMatrix(k, .ColIndex("LastDateH")) = LastDateH.value
.TextMatrix(k, .ColIndex("VacBalance")) = val(TXtVacBalance.text)
.TextMatrix(k, .ColIndex("VacWithoutSal")) = val(TxtVacWithoutSal.text)
.TextMatrix(k, .ColIndex("Abcence")) = val(TxtAbcence.text)
k = k + 1
Loop
'DcboEmpName.BoundText = 0
'TXtVacBalance.text = 0
'TxtVacWithoutSal.text = 0
'TxtAbcence.text = ""
End With
End Sub
Sub filgrid2(Optional index As Integer = 0)
Dim k As Integer
Dim i As Integer
Dim sql As String
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset
sql = "SELECT     dbo.Contract.Contract_date, dbo.Contract.DateH, dbo.TblEmployee.*"
sql = sql & " FROM         dbo.TblEmployee LEFT OUTER JOIN"
sql = sql & "                      dbo.Contract ON dbo.TblEmployee.Emp_ID = dbo.Contract.Emp_id"

        If SystemOptions.OpeningEmployeeShowAll = True Then
sql = sql & "                      WHERE     (1=1)"
Else
sql = sql & "                      WHERE     (TblEmployee.BignDateWork IS NULL)"
End If

If index = 1 Then
sql = sql & " and     (TblEmployee.BranchId =" & val(DcbSelBranch.BoundText) & ")  "
ElseIf index = 2 Then
sql = sql & " and     (TblEmployee.DepartmentID  =" & val(DcbDepartment.BoundText) & ")  "
End If
Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs1.RecordCount > 0 Then
With Grid
k = .rows - 1
.rows = .rows + Rs1.RecordCount
Rs1.MoveFirst
Do While k < (.rows - 1)
.TextMatrix(k, .ColIndex("Ser")) = k
.TextMatrix(k, .ColIndex("EmpID")) = IIf(IsNull(Rs1("Emp_ID").value), "", Rs1("Emp_ID").value)
.TextMatrix(k, .ColIndex("Fullcode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs1("Emp_Name").value), "", Rs1("Emp_Name").value)
Else
.TextMatrix(k, .ColIndex("name")) = IIf(IsNull(Rs1("Emp_Namee").value), "", Rs1("Emp_Namee").value)
End If
.TextMatrix(k, .ColIndex("BeginDate")) = IIf(IsNull(Rs1("Contract_date").value), BeginDate.value, Rs1("Contract_date").value)
.TextMatrix(k, .ColIndex("BeginDateH")) = IIf(IsNull(Rs1("DateH").value), BeginDateH.value, Rs1("DateH").value)
.TextMatrix(k, .ColIndex("LastDate")) = LastDate.value
.TextMatrix(k, .ColIndex("LastDateH")) = LastDateH.value
.TextMatrix(k, .ColIndex("VacBalance")) = val(TXtVacBalance.text)
.TextMatrix(k, .ColIndex("VacWithoutSal")) = val(TxtVacWithoutSal.text)
.TextMatrix(k, .ColIndex("Abcence")) = val(TxtAbcence.text)
k = k + 1
Rs1.MoveNext
Loop
'DcboEmpName.BoundText = 0
'TXtVacBalance.text = 0
'TxtVacWithoutSal.text = 0
'TxtAbcence.text = ""
End With
End If
End Sub

Private Sub ISButton3_Click()
If Me.TxtModFlg.text <> "R" Then
Dim sql As String
Dim StrSQL As String
 On Error Resume Next
    With Me.Grid
        If .row <= 0 Then Exit Sub
        If CheckEmp(val(.TextMatrix(.row, .ColIndex("EmpID")))) = False Then
        sql = "update TblEmployee set   BignDateWork = null,IssueDateH=null  where Emp_ID =" & val(.TextMatrix(.row, .ColIndex("EmpID"))) & ""
         Cn.Execute sql
        sql = "update TblEmployee set   lastHolidaydate = null,lastHolidaydateH=null where Emp_ID =" & val(.TextMatrix(.row, .ColIndex("EmpID"))) & ""
         Cn.Execute sql
         StrSQL = "Delete From TblInstalVacationDet Where InslVaID=" & val(TxtSerial1.text) & " and EmpID=" & val(.TextMatrix(.row, .ColIndex("EmpID"))) & ""
         Cn.Execute StrSQL, , adExecuteNoRecords
        .RemoveItem .row
        Else
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–› ·«— »«ÿÂ »«·„»«‘—« "
        Else
        MsgBox "You can not delete this employee exists in start a work "
        End If
        Exit Sub
        End If
    End With
 End If
End Sub
Function CheckEmp(Optional EmpID As Double) As Boolean
Dim Rs9 As ADODB.Recordset
Set Rs9 = New ADODB.Recordset
Dim sql As String
sql = " SELECT     Emp_ID"
sql = sql & " From dbo.TblEmbarkation"
sql = sql & " Where (Emp_id = " & EmpID & ")"
Rs9.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs9.RecordCount > 0 Then
CheckEmp = True
Else
CheckEmp = False
End If
End Function
Private Sub ISButton4_Click()
If Me.TxtModFlg.text <> "R" Then
Dim k As Integer
Dim i As Integer
Dim sql As String
Dim StrSQL As String
On Error Resume Next
With Grid
k = .rows - 1
For i = 1 To .rows - 1
If CheckEmp(val(.TextMatrix(k, .ColIndex("EmpID")))) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–› ·«— »«ÿÂ »«·„»«‘—« "
        Else
        MsgBox "You can not delete this employee exists in start a work "
        End If
Exit Sub
End If
If k <= 0 Then Exit Sub
   sql = "update TblEmployee set   BignDateWork = null,IssueDateH=null  where Emp_ID =" & val(.TextMatrix(k, .ColIndex("EmpID"))) & ""
   Cn.Execute sql
   sql = "update TblEmployee set   lastHolidaydate = null,lastHolidaydateH=null where Emp_ID =" & val(.TextMatrix(k, .ColIndex("EmpID"))) & ""
   Cn.Execute sql
   StrSQL = "Delete From TblInstalVacationDet Where InslVaID=" & val(TxtSerial1.text) & " and EmpID=" & val(.TextMatrix(k, .ColIndex("EmpID"))) & ""
         Cn.Execute StrSQL, , adExecuteNoRecords
.RemoveItem k
k = k - 1
Next i
End With
'Me.Grid.Clear flexClearScrollable, flexClearEverything
'cleargriid

 '    Me.Grid.Clear flexClearScrollable, flexClearEverything
  '   Grid.Rows = 2
 End If
End Sub



Private Sub ISButton5_Click()
print_report
End Sub

Private Sub ISButton8_Click()
Load FrmInstalVacationSearch
FrmInstalVacationSearch.show
End Sub

Private Sub LastDate_Change()
 If Me.TxtModFlg.text <> "R" Then
              LastDateH.value = ToHijriDate(LastDate.value)
   End If
End Sub

Private Sub LastDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
  VBA.Calendar = vbCalGreg
            LastDate.value = ToGregorianDate(LastDateH.value)
 End If
End Sub
  
Private Sub Opt_Click(index As Integer)
Dim My_SQL As String
     
        If SystemOptions.OpeningEmployeeShowAll = True Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                                    My_SQL = "SELECT Emp_ID,Emp_Name From TblEmployee    "
                                    My_SQL = My_SQL + " Order By Emp_Name ASC"
                        Else
                                    My_SQL = "SELECT Emp_ID,Emp_Namee From TblEmployee    "
                                    My_SQL = My_SQL + " Order By Emp_Namee ASC"
                        End If
    Else
    
                      If SystemOptions.UserInterface = ArabicInterface Then
                                 My_SQL = "SELECT Emp_ID,Emp_Name From TblEmployee WHERE     (BignDateWork IS NULL)   "
                                 My_SQL = My_SQL + " Order By Emp_Name ASC"
                       Else
                                 My_SQL = "SELECT Emp_ID,Emp_Namee From TblEmployee WHERE     (BignDateWork IS NULL)   "
                                 My_SQL = My_SQL + " Order By Emp_Namee ASC"
                       End If
                    
    
    End If
    


 
    fill_combo DcboEmpName, My_SQL
If index = 0 Then
DcboEmpName.Enabled = True
TxtSearchCode.Enabled = True
DcbSelBranch.Enabled = False
DcbDepartment.Enabled = False
ElseIf index = 1 Then
DcboEmpName.Enabled = False
TxtSearchCode.Enabled = False
DcbSelBranch.Enabled = True
DcbDepartment.Enabled = False
ElseIf index = 2 Then
DcboEmpName.Enabled = False
TxtSearchCode.Enabled = False
DcbSelBranch.Enabled = False
DcbDepartment.Enabled = True
ElseIf index = 3 Then
DcboEmpName.Enabled = False
TxtSearchCode.Enabled = False
DcbSelBranch.Enabled = False
DcbDepartment.Enabled = False
End If
End Sub

Private Sub Txt_DateHigri_LostFocus()
If Me.TxtModFlg.text <> "R" Then
  VBA.Calendar = vbCalGreg
            XPDtbTrans.value = ToGregorianDate(Txt_DateHigri.value)
 End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub


Private Sub TXtVacBalance_KeyPress(KeyAscii As Integer)
       On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  TxtVacWithoutSal.SetFocus
  End If
ErrTrap:
End Sub

Private Sub TxtVacWithoutSal_KeyPress(KeyAscii As Integer)
       On Error GoTo ErrTrap
  If KeyAscii = 13 Then
  TxtAbcence.SetFocus
  End If
ErrTrap:
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
   ' On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
      If dcBranch.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            dcBranch.SetFocus
            Exit Sub
            Else
            MsgBox "Write Arabic Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
            dcBranch.SetFocus
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
    StrRecID = new_id("TblInstalVacation", "ID", "")
    Me.TxtSerial1.text = StrRecID
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
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
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1
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
     BtnLast_Click
End Sub
' delet sub
Private Sub btnDelete_Click()
Dim sql As String
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim i As Integer
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
  With Grid
  For i = 1 To .rows - 1
  If CheckEmp(val(.TextMatrix(i, .ColIndex("EmpID")))) = True Then
      If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «·Õ–› ·«— »«ÿÂ »«·„»«‘—« "
        Else
        MsgBox "You can not delete this employee exists in start a work "
        End If
  Exit Sub
  End If
  Next i
  For i = 1 To .rows - 1
     sql = "update TblEmployee set   BignDateWork = null,IssueDateH=null  where Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
   Cn.Execute sql
   sql = "update TblEmployee set   lastHolidaydate = null,lastHolidaydateH=null where Emp_ID =" & val(.TextMatrix(i, .ColIndex("EmpID"))) & ""
   Cn.Execute sql
  Next i
  End With
  
    If X = vbNo Then Exit Sub
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
                RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                RsSavRec.delete
               '''''''''''''''''''''''''''''''
                 StrSQL = "Delete From TblInstalVacationDet Where InslVaID='" & val(TxtSerial1.text) & "'"
                 Cn.Execute StrSQL, , adExecuteNoRecords
                  StrSQL = "Delete From TblInforVacatiom Where InstVacaID='" & val(TxtSerial1.text) & "'"
                  Cn.Execute StrSQL, , adExecuteNoRecords
                    StrSQL = "Delete From tblVacationData Where InstVacaID='" & val(TxtSerial1.text) & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Deletion Process Success ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
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
    RsSavRec.MoveFirst
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
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
    RsSavRec.MoveLast
    cleargriid
    FiLLTXT
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    Dim Msg As String
    Dim i As Integer
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.text <> "" Then
          TxtModFlg = "E"
        Grid.rows = Grid.rows + 1
        Me.DCboUserName.BoundText = user_id
      '  Me.Dcbranch.BoundText = branch_id
        Frm2.Enabled = True
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
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
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    cleargriid
    TxtModFlg.text = "N"
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = branch_id
    dcBranch.SetFocus
     Me.Grid.Clear flexClearScrollable, flexClearEverything
     Grid.rows = 2
     Dim FirstPeriodDateInthisYear As Date
         getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear

    Me.XPDtbTrans.value = FirstPeriodDateInthisYear


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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
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
  MySQL = " SELECT     dbo.TblInstalVacation.ID, dbo.TblInstalVacation.RecordM, dbo.TblInstalVacation.RecordH, dbo.TblInstalVacation.BranchID, dbo.TblBranchesData.branch_name, "
  MySQL = MySQL & "                    dbo.TblBranchesData.branch_namee, dbo.TblInstalVacation.TypeSelect, dbo.TblInstalVacation.BranchSelectID, TblBranchesData_1.branch_name AS Selbranch_name,"
  MySQL = MySQL & "                     TblBranchesData_1.branch_namee AS Selbranch_nameE, dbo.TblInstalVacation.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
  MySQL = MySQL & "                     dbo.TblEmployee.Emp_Namee, dbo.TblInstalVacation.DeptID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
  MySQL = MySQL & "                     dbo.TblInstalVacation.VacBalance, dbo.TblInstalVacation.VacWithoutSal, dbo.TblInstalVacation.Abcence, dbo.TblInstalVacation.BeginDate,"
  MySQL = MySQL & "                     dbo.TblInstalVacation.BeginDateH, dbo.TblInstalVacation.LastDate, dbo.TblInstalVacation.LastDateH, dbo.TblInstalVacationDet.BeginDate AS BeginDateDet,"
  MySQL = MySQL & "                     dbo.TblInstalVacationDet.LastDate AS LastDateDet, dbo.TblInstalVacationDet.VacBalance AS VacBalanceDet,"
  MySQL = MySQL & "                     dbo.TblInstalVacationDet.VacWithoutSal AS VacWithoutSalDet, dbo.TblInstalVacationDet.Abcence AS AbcenceDet,"
  MySQL = MySQL & "                     dbo.TblInstalVacationDet.BeginDateH AS BeginDateHDet, dbo.TblInstalVacationDet.LastDateH AS LastDateHDet, TblEmployee_1.Emp_Name AS Emp_NameDet,"
  MySQL = MySQL & "                     TblEmployee_1.Fullcode AS FullcodeDet, TblEmployee_1.Emp_Namee AS Expr12Emp_NameeDet"
  MySQL = MySQL & "  FROM         dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblInstalVacationDet ON TblEmployee_1.Emp_ID = dbo.TblInstalVacationDet.EmpID RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblInstalVacation INNER JOIN"
  MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblInstalVacation.BranchID = dbo.TblBranchesData.branch_id ON"
  MySQL = MySQL & "                     dbo.TblInstalVacationDet.InslVaID = dbo.TblInstalVacation.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblEmpDepartments ON dbo.TblInstalVacation.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblEmployee ON dbo.TblInstalVacation.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblBranchesData TblBranchesData_1 ON dbo.TblInstalVacation.BranchSelectID = TblBranchesData_1.branch_id"
  MySQL = MySQL & "       Where (dbo.TblInstalVacation.id =" & val(TxtSerial1.text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepInstalVacation.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepInstalVacationE.rpt"
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

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
       
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
'Information for camand
'++++++++++++++++++++++++++++++++++++++
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
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
            Sendkeys "{TAB}"
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
'On Error GoTo ErrTrap
   ' form name
   
      Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic


    Me.Caption = "Opening Balances"
    
    ' labell name
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(4).Caption = "Trans ID"
    Me.lbl(2).Caption = "Date"
   
    Me.Label3.Caption = "Branch"
    opt(0).Caption = "Select Emp"
    opt(0).RightToLeft = False
    
     opt(1).Caption = "Select Branch"
     opt(1).RightToLeft = False
     opt(2).Caption = "Select Dept"
     opt(2).RightToLeft = False
     opt(3).Caption = "Select All"
     opt(3).RightToLeft = False
    lbl(0).Caption = "Start Date"
    lbl(1).Caption = "Last Date"
    lbl(3).Caption = "Balances Vaca"
    lbl(5).Caption = "Unpaid Vaca "
    lbl(6).Caption = "Absence"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
ISButton2.Caption = "Add"

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
        .TextMatrix(0, .ColIndex("name")) = "Employee Name"
         .TextMatrix(0, .ColIndex("BeginDate")) = "Start Date"
        .TextMatrix(0, .ColIndex("BeginDateH")) = "Start Date"
        .TextMatrix(0, .ColIndex("LastDate")) = "Last Date"
         .TextMatrix(0, .ColIndex("LastDateH")) = "Last Date"
        .TextMatrix(0, .ColIndex("VacBalance")) = "Balances Vacation"
        .TextMatrix(0, .ColIndex("VacWithoutSal")) = "Unpaid Vacation "
        .TextMatrix(0, .ColIndex("Abcence")) = "Abcence"
        
    End With
ErrTrap:
End Sub
Private Sub cleargriid()
Me.Grid.rows = 1
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblInstalVacation"
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






