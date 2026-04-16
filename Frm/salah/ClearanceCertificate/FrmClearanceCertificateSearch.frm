VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmClearanceCerificateSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ‘Â«œ… ≈Œ·«¡ ÿ—›"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10335
   Icon            =   "FrmClearanceCertificateSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5340
   ScaleWidth      =   10335
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   5055
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   -240
      Width           =   10335
      Begin VB.TextBox TxtCus_Name 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   3600
         Width           =   4740
      End
      Begin VB.TextBox TxtTypeRequest 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   3120
         Width           =   1455
      End
      Begin VB.TextBox TxtTelephone 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox ItemRequest 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4275
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   62
         Top             =   4680
         Width           =   4815
      End
      Begin VB.TextBox TxtIBN 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   60
         Top             =   4560
         Width           =   2895
      End
      Begin VB.TextBox TxtCust_Mobile 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   7635
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   3960
         Width           =   1455
      End
      Begin VB.TextBox TxtCity 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   4080
         Width           =   2895
      End
      Begin VB.TextBox TxtCusID 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox TxtHome_Tel 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   7635
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   4320
         Width           =   1455
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   14280
         TabIndex        =   46
         Top             =   3600
         Width           =   810
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ "
         Height          =   1155
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2880
         Width           =   4095
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   330
            Left            =   2010
            TabIndex        =   40
            Top             =   270
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   60817411
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   330
            Left            =   90
            TabIndex        =   41
            Top             =   270
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   60817411
            CurrentDate     =   38887
         End
         Begin Dynamic_Byte.NourHijriCal BrithDateH 
            Height          =   330
            Left            =   90
            TabIndex        =   56
            Top             =   720
            Width           =   1590
            _extentx        =   2805
            _extenty        =   582
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
            Height          =   330
            Left            =   2010
            TabIndex        =   57
            Top             =   720
            Width           =   1590
            _extentx        =   2805
            _extenty        =   582
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   12
            Left            =   1695
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   420
            Width           =   255
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   11
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   420
            Width           =   300
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   645
         Index           =   0
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2880
         Width           =   3555
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   9
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   300
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   8
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   405
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2625
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   10155
         _cx             =   17912
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
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
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmClearanceCertificateSearch.frx":038A
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
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   10320
         TabIndex        =   47
         Top             =   3600
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         BoundColumn     =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”·⁄… «·„ÿ·Ê»…"
         Height          =   285
         Index           =   22
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   4680
         Width           =   1965
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Õ”«» «·»‰ﬂÌ"
         Height          =   285
         Index           =   20
         Left            =   2640
         TabIndex        =   61
         Top             =   4560
         Width           =   1965
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÊ«· «·⁄„Ì·"
         Height          =   285
         Index           =   19
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   3960
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‰ÿﬁ…"
         Height          =   285
         Index           =   18
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   4080
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÂÊÌ…"
         Height          =   285
         Index           =   16
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Â« › «·⁄„·"
         Height          =   285
         Index           =   17
         Left            =   6135
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   3960
         Width           =   1965
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Â« › «·„‰“·"
         Height          =   285
         Index           =   21
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   4320
         Width           =   1965
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì·"
         Height          =   285
         Index           =   14
         Left            =   8850
         TabIndex        =   48
         Top             =   3660
         Width           =   1965
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ÿ·»"
         Height          =   285
         Index           =   13
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   3120
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   5055
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   -240
      Width           =   10335
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·⁄„·Ì…"
         Height          =   645
         Index           =   2
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2880
         Width           =   3555
         Begin VB.TextBox TxtIDTO 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   180
            Width           =   915
         End
         Begin VB.TextBox TxtIDFrom 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   180
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   6
            Left            =   1020
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   5
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   240
            Width           =   300
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   645
         Index           =   1
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2880
         Width           =   4095
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   2010
            TabIndex        =   24
            Top             =   270
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   60817411
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   90
            TabIndex        =   25
            Top             =   270
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   60817411
            CurrentDate     =   38887
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   4
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   330
            Width           =   180
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   3
            Left            =   1695
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   300
            Width           =   255
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   2625
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10155
         _cx             =   17912
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14871017
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
         FormatString    =   $"FrmClearanceCertificateSearch.frx":0594
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
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   495
         Left            =   4920
         TabIndex        =   9
         Top             =   4440
         Width           =   5295
         _Version        =   786432
         _ExtentX        =   9340
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RDVacation 
            Height          =   255
            Left            =   3960
            TabIndex        =   10
            Top             =   120
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "≈Ã«“…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RDTransfer 
            Height          =   255
            Left            =   3000
            TabIndex        =   11
            Top             =   120
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "‰ﬁ·"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RDFinalExit 
            Height          =   255
            Left            =   1560
            TabIndex        =   12
            Top             =   120
            Width           =   1335
            _Version        =   786432
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Œ—ÊÃ ‰Â«∆Ì"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RDAll 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   120
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   975
         Left            =   0
         TabIndex        =   14
         Top             =   3480
         Width           =   10215
         _Version        =   786432
         _ExtentX        =   18018
         _ExtentY        =   1720
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin MSDataListLib.DataCombo DCEmp_Name 
            Height          =   315
            Left            =   5280
            TabIndex        =   15
            Top             =   240
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   "DCEmp_Name"
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpDepartments 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   210
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbNatinalty 
            Height          =   315
            Left            =   5280
            TabIndex        =   17
            Top             =   570
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboJobsType1 
            Height          =   315
            Left            =   120
            TabIndex        =   18
            Top             =   570
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   315
            Index           =   0
            Left            =   8940
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—…"
            Height          =   285
            Index           =   15
            Left            =   4080
            TabIndex        =   21
            Top             =   210
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ã‰”Ì…"
            Height          =   285
            Index           =   7
            Left            =   9270
            TabIndex        =   20
            Top             =   570
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÊŸÌ›…"
            Height          =   285
            Index           =   24
            Left            =   4080
            TabIndex        =   19
            Top             =   570
            Width           =   645
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   9330
      TabIndex        =   0
      Top             =   4920
      Width           =   765
      _ExtentX        =   1349
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
      Left            =   8490
      TabIndex        =   1
      Top             =   4920
      Width           =   795
      _ExtentX        =   1402
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
      Left            =   7710
      TabIndex        =   2
      Top             =   4920
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSDataListLib.DataCombo DcboJobsType 
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   -480
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   5970
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4980
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4980
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4860
      Width           =   2535
   End
End
Attribute VB_Name = "FrmClearanceCerificateSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DCboSearch As clsDCboSearch
Public ind As Integer

Private Sub BrithDateH_LostFocus()
DTPicker2.value = ToGregorianDate(BrithDateH.value)
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
        If ind = 0 Then
            GetData
         ElseIf ind = 1 Then
         GetDataInstament
         End If

        Case 1
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
DTPicker2.value = ""
DTPicker1.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub



'Private Sub DBCboClientName_Click(Area As Integer)
'DBCboClientName_Change
'End Sub

Private Sub DTPicker1_Change()
If Not (IsNull(DTPicker1.value)) Then
         NourHijriCal1.value = ToHijriDate(DTPicker1.value)
 End If
End Sub

Private Sub DTPicker2_Change()
If Not (IsNull(DTPicker2.value)) Then
         BrithDateH.value = ToHijriDate(DTPicker2.value)
 End If
End Sub

Private Sub fg_Click()
If ind = 0 Then
    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("AdvanceID"))) = 0 Then
            Exit Sub
        End If

       
                FrmClearanceCerTifcate.Retrive val(.TextMatrix(.Row, .ColIndex("AdvanceID")))
           
       

    End With
End If
End Sub
Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  

'Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "EmpName"
lbl(15).Caption = "Dept"
lbl(7).Caption = "Nationality"
lbl(24).Caption = "Position"
Fra(1).Caption = "Date Registration"
Fra(2).Caption = "Process No"
lbl(2).Caption = "Total"
RDVacation.RightToLeft = False
RDVacation.Caption = "Vacation"
RDTransfer.RightToLeft = False
RDTransfer.Caption = "Transfer"
RDFinalExit.RightToLeft = False
RDFinalExit.Caption = "Final Exit"
RdAll.RightToLeft = False
lbl(20).Caption = "IBN"
RdAll.Caption = "All"
lbl(18).Caption = "City"
lbl(11).Caption = "From"
lbl(12).Caption = "To"
lbl(13).Caption = "Type"
Fra(3).Caption = "Date"
Fra(0).Caption = "Request No."
lbl(9).Caption = "From"
lbl(8).Caption = "To"
lbl(14).Caption = "Customer"
lbl(19).Caption = "Mobile"
lbl(17).Caption = "Work Tele."
lbl(21).Caption = "Home Tele"
lbl(16).Caption = "ID"
lbl(22).Caption = "Goods"
     With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("AdvanceID")) = "Code"
        .TextMatrix(0, .ColIndex("AdvanceDate")) = "Date"
         .TextMatrix(0, .ColIndex("Emp_Name")) = "Emp_Name"
        .TextMatrix(0, .ColIndex("nationality")) = "Nationality"
       .TextMatrix(0, .ColIndex("job")) = "Position"
        .TextMatrix(0, .ColIndex("Dept")) = "Dept"
    End With
       With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "Serila"
        .TextMatrix(0, .ColIndex("ID")) = "No"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("RecordDateH")) = "Date"
        .TextMatrix(0, .ColIndex("TypeRequest")) = "Type"
       .TextMatrix(0, .ColIndex("ItemRequest")) = "Goods"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer"
        .TextMatrix(0, .ColIndex("Cust_Mobile")) = "Mobile"
        .TextMatrix(0, .ColIndex("Telephone")) = "Work_Tel"
        .TextMatrix(0, .ColIndex("Home_Tel")) = "Home_Tel"
        .TextMatrix(0, .ColIndex("CusID")) = "ID"
        .TextMatrix(0, .ColIndex("City")) = "City"
        .TextMatrix(0, .ColIndex("IBN")) = "IBN"
    End With
End Sub

Private Sub NourHijriCal1_LostFocus()
DTPicker1.value = ToGregorianDate(NourHijriCal1.value)
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub
Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
    If ind = 0 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    Me.Caption = "«·»ÕÀ ⁄‰ ‘Â«œ… «Œ·«¡ ÿ—›"
    Else
    Me.Caption = "Search Clearance Certifcate"
    End If
    ElseIf ind = 1 Then
    If SystemOptions.UserInterface = ArabicInterface Then
    Me.Caption = "«·»ÕÀ ⁄‰ ÿ·» ‘—«¡ ”·⁄… »«· ﬁ”Ìÿ"
    Else
    Me.Caption = "Search Buy Goods in Installments Request"
    End If
    End If
End Sub


'Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
'  Dim CUSTID As Integer
'
'    If KeyAscii = vbKeyReturn Then
'        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
'        DBCboClientName.BoundText = CUSTID
'    End If
'
'End Sub
'Private Sub DBCboClientName_Change()
'If val(DBCboClientName.BoundText) <> 0 Then
'    Dim FullCode As String
'    GetCustomersDetail val(DBCboClientName.BoundText), , FullCode
'    TxtSearchCode.text = FullCode
'    End If
'    End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
If ind = 0 Then
Frame2.Visible = False
Frame1.Visible = True
ElseIf ind = 1 Then
Frame2.Visible = True
Frame1.Visible = False
End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DCEmp_Name
    Set DCboSearch = New clsDCboSearch
    Set DCboSearch.Client = Me.DCEmp_Name
    Dcombos.GetEmpDepartments Me.DcboEmpDepartments
    Dcombos.GetCustomersSuppliers 1, DBCboClientName
    Dcombos.GetEmpJobsTypes Me.DcboJobsType1

    'Dcombos.GetEmpGrades Me.DcboSpecifications
    Dcombos.GetEmployeesNationlity Me.DcbNatinalty
  RdAll.value = True
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 SetDtpickerDate DTPicker1
 SetDtpickerDate DTPicker2
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
       ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Private Sub GetData()
    Dim strSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
strSQL = "SELECT     dbo.TblClearanceCertificate.ID, dbo.TblClearanceCertificate.RecordDate, dbo.TblClearanceCertificate.LeavingDate, dbo.TblClearanceCertificate.Posted,"
 strSQL = strSQL & "                     dbo.TblClearanceCertificate.UserID, dbo.TblClearanceCertificate.nationality, dbo.TblClearanceCertificate.vacation, dbo.TblClearanceCertificate.Transfer,"
 strSQL = strSQL & "                      dbo.TblClearanceCertificate.FinalExit, dbo.TblClearanceCertificate.EmpID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
 strSQL = strSQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality AS natinalityc,"
 strSQL = strSQL & "                      dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee2,"
 strSQL = strSQL & "                      dbo.TblClearanceCertificate.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblClearanceCertificate.GradID,"
  strSQL = strSQL & "                     dbo.TblEmpGrades.Lowsalary, dbo.TblEmpGrades.HighSalary, dbo.TblEmpGrades.name, dbo.TblEmpGrades.namee, dbo.TblClearanceCertificate.JobID,"
 strSQL = strSQL & "                      dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblClearanceCertificate.DeptID, dbo.TblEmpDepartments.DepartmentName,"
 strSQL = strSQL & "                      dbo.TblEmpDepartments.DepartmentNamee"
 strSQL = strSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
  strSQL = strSQL & "                     dbo.TblEmpGrades RIGHT OUTER JOIN"
   strSQL = strSQL & "                    dbo.TblClearanceCertificate INNER JOIN"
 strSQL = strSQL & "                      dbo.TblBranchesData ON dbo.TblClearanceCertificate.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 strSQL = strSQL & "                      dbo.TblEmpDepartments ON dbo.TblClearanceCertificate.DeptID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
 strSQL = strSQL & "                      dbo.TblEmpJobsTypes ON dbo.TblClearanceCertificate.JobID = dbo.TblEmpJobsTypes.JobTypeID ON"
 strSQL = strSQL & "                      dbo.TblEmpGrades.gradeid = dbo.TblClearanceCertificate.GradID ON dbo.TblEmployee.Emp_ID = dbo.TblClearanceCertificate.EmpID"

    BolBegine = False
    StrWhere = ""
 If Me.RDVacation.value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.vacation = 1 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.vacation = 1 "
        End If
    End If
    If Me.RDTransfer.value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.Transfer = 1 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.Transfer = 1 "
        End If
    End If
    If Me.RDFinalExit.value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.FinalExit = 1 "
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.FinalExit = 1 "
        End If
    End If
     
    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If

    If Me.DCEmp_Name.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.EmpID=" & Me.DCEmp_Name.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.EmpID=" & Me.DCEmp_Name.BoundText & ""
        End If
    End If
 If Me.DcbNatinalty.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.nationality='" & Me.DcbNatinalty.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.nationality='" & Me.DcbNatinalty.text & "'"
        End If
    End If
     If Me.DcboEmpDepartments.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.DeptID=" & Me.DcboEmpDepartments.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.DeptID=" & Me.DcboEmpDepartments.BoundText & ""
        End If
    End If
     If Me.DcboJobsType1.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.JobID=" & Me.DcboJobsType1.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.JobID=" & Me.DcboJobsType1.BoundText & ""
        End If
    End If
  

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblClearanceCertificate.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblClearanceCertificate.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    strSQL = strSQL & StrWhere
    strSQL = strSQL & " Order By dbo.TblClearanceCertificate.ID"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("AdvanceID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("AdvanceDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("nationality")) = IIf(IsNull(rs("nationality").value), "", rs("nationality").value)
                .TextMatrix(i, .ColIndex("job")) = IIf(IsNull(rs("JobTypeName").value), "", rs("JobTypeName").value)
                .TextMatrix(i, .ColIndex("Dept")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
           ' Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub
Private Sub GetDataInstament()
    Dim strSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
strSQL = "SELECT     dbo.TblInstallmentsReq.ID, dbo.TblInstallmentsReq.RecordDate, dbo.TblInstallmentsReq.RecordDateH, dbo.TblInstallmentsReq.BranchID, "
strSQL = strSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblInstallmentsReq.TypeRequest, dbo.TblInstallmentsReq.Cust_Mobile,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.ItemRequest, dbo.TblInstallmentsReq.Cus_ID, dbo.TblInstallmentsReq.BrithDateH, dbo.TblInstallmentsReq.BrithDate,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Company, dbo.TblInstallmentsReq.City, dbo.TblInstallmentsReq.CusID, dbo.TblInstallmentsReq.ExpDate, dbo.TblInstallmentsReq.PlaceID,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Salary, dbo.TblInstallmentsReq.BankName, dbo.TblInstallmentsReq.Salary_Acc, dbo.TblInstallmentsReq.Telephone,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Ext, dbo.TblInstallmentsReq.Home_Tel, dbo.TblInstallmentsReq.Street, dbo.TblInstallmentsReq.District, dbo.TblInstallmentsReq.Remarks,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.[No], dbo.TblInstallmentsReq.IBN, dbo.TblInstallmentsReq.Total_liab, dbo.TblInstallmentsReq.TypeAccept, dbo.TblInstallmentsReq.GurName,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.GurBrithDate, dbo.TblInstallmentsReq.GurBrithDateH, dbo.TblInstallmentsReq.GurCompany, dbo.TblInstallmentsReq.GurSalary,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.GurTotal_liab, dbo.TblInstallmentsReq.Relation1, dbo.TblInstallmentsReq.Relation2, dbo.TblInstallmentsReq.Adress1,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Adress2, dbo.TblInstallmentsReq.GurType_liab, dbo.TblInstallmentsReq.DirectManager, dbo.TblInstallmentsReq.GurTele,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.GurExt, dbo.TblInstallmentsReq.NameOffice, dbo.TblInstallmentsReq.NameAdmin_Mobile, dbo.TblInstallmentsReq.Attch0,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Attch1, dbo.TblInstallmentsReq.Attch2, dbo.TblInstallmentsReq.Attch3, dbo.TblInstallmentsReq.Attch4, dbo.TblInstallmentsReq.Attch5,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Attch6, dbo.TblInstallmentsReq.Attch7, dbo.TblInstallmentsReq.Attch8, dbo.TblInstallmentsReq.Attch9, dbo.TblInstallmentsReq.Attch10,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Attch11, dbo.TblInstallmentsReq.Attch12, dbo.TblInstallmentsReq.Attch13, dbo.TblInstallmentsReq.Attch14, dbo.TblInstallmentsReq.Attch15,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Attch16, dbo.TblInstallmentsReq.Attch17, dbo.TblInstallmentsReq.Attch18, dbo.TblInstallmentsReq.Accept, dbo.TblInstallmentsReq.Attch19,"
strSQL = strSQL & "                      dbo.TblInstallmentsReq.Cus_Name"
strSQL = strSQL & " FROM         dbo.TblInstallmentsReq LEFT OUTER JOIN"
strSQL = strSQL & "                      dbo.TblBranchesData ON dbo.TblInstallmentsReq.BranchID = dbo.TblBranchesData.branch_id"
    BolBegine = False
    StrWhere = ""

     
    If val(Me.Text2.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.ID >=" & val(Me.Text2.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.ID >=" & val(Me.Text2.text) & ""
        End If
    End If

    If val(Me.Text1.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.ID <=" & val(Me.Text1.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.ID <=" & val(Me.Text1.text) & ""
        End If
    End If

    'If Me.DBCboClientName.text <> "" And val(DBCboClientName.BoundText) <> 0 Then
    '    If BolBegine = True Then
    '        StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.Cus_ID=" & val(Me.DBCboClientName.BoundText) & ""
    '    Else
    '        BolBegine = True
    '        StrWhere = " Where dbo.TblInstallmentsReq.Cus_ID=" & val(Me.DBCboClientName.BoundText) & ""
    '    End If
    'End If
 If Me.TxtCus_Name.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.Cus_Name like '%" & Me.TxtCus_Name.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.Cus_Name  like '%" & Me.TxtCus_Name.text & "%'"
        End If
    End If
     If Me.TxtTypeRequest.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.TypeRequest like '%" & Me.TxtTypeRequest.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.TypeRequest  like '%" & Me.TxtTypeRequest.text & "%'"
        End If
    End If
     If Me.ItemRequest.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.ItemRequest like '%" & Me.ItemRequest.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.ItemRequest like '%" & Me.ItemRequest.text & "%'"
        End If
    End If
     If Me.TxtCust_Mobile.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.Cust_Mobile='" & Me.TxtCust_Mobile.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.Cust_Mobile='" & Me.TxtCust_Mobile.text & "'"
        End If
    End If
    If Me.TxtHome_Tel.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.Home_Tel='" & Me.TxtHome_Tel.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.Home_Tel='" & Me.TxtHome_Tel.text & "'"
        End If
    End If
    
    If Me.TxtTelephone.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.Telephone='" & Me.TxtTelephone.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.Telephone='" & Me.TxtTelephone.text & "'"
        End If
    End If
        If Me.TxtCusID.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.CusID='" & Me.TxtCusID.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.CusID='" & Me.TxtCusID.text & "'"
        End If
    End If
          If Me.TxtCity.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.City='" & Me.TxtCity.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.City='" & Me.TxtCity.text & "'"
        End If
    End If
              If Me.TxtIBN.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.IBN='" & Me.TxtIBN.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.IBN='" & Me.TxtIBN.text & "'"
        End If
    End If
    If Not IsNull(Me.DTPicker1.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.RecordDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.RecordDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DTPicker2.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblInstallmentsReq.RecordDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblInstallmentsReq.RecordDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
        End If
    End If

    '-----------------------------------

    strSQL = strSQL & StrWhere
    strSQL = strSQL & " Order By dbo.TblInstallmentsReq.ID"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
      Else
      Msg = "No Data"
      End If
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                .TextMatrix(i, .ColIndex("IBN")) = IIf(IsNull(rs("IBN").value), "", rs("IBN").value)
                .TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(rs("RecordDateH").value), ToHijriDate(Date), rs("RecordDateH").value)
                .TextMatrix(i, .ColIndex("TypeRequest")) = IIf(IsNull(rs("TypeRequest").value), "", rs("TypeRequest").value)
                .TextMatrix(i, .ColIndex("Cust_Mobile")) = IIf(IsNull(rs("Cust_Mobile").value), "", rs("Cust_Mobile").value)
                .TextMatrix(i, .ColIndex("ItemRequest")) = IIf(IsNull(rs("ItemRequest").value), "", rs("ItemRequest").value)
                .TextMatrix(i, .ColIndex("City")) = IIf(IsNull(rs("City").value), "", rs("City").value)
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
                .TextMatrix(i, .ColIndex("Home_Tel")) = IIf(IsNull(rs("Home_Tel").value), "", rs("Home_Tel").value)
              
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("Cus_Name").value), "", rs("Cus_Name").value)
               
                
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
           ' Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub


Private Sub VSFlexGrid1_Click()
If ind = 1 Then
With VSFlexGrid1
FrmBuyGoodsInst.FindRec val(.TextMatrix(.Row, .ColIndex("ID")))
End With
End If
End Sub
