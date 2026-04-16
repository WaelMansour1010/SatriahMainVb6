VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmExpensesInvestment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19545
   Icon            =   "FrmExpensesInvestment.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   19545
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
      Left            =   19680
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmExpensesInvestment.frx":6852
      Left            =   19560
      List            =   "FrmExpensesInvestment.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   28
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
      TabIndex        =   22
      Top             =   0
      Width           =   19545
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox tXTRootAccount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
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
         ButtonImage     =   "FrmExpensesInvestment.frx":687B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   24
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
         ButtonImage     =   "FrmExpensesInvestment.frx":6C15
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   25
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
         ButtonImage     =   "FrmExpensesInvestment.frx":6FAF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   26
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
         ButtonImage     =   "FrmExpensesInvestment.frx":7349
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„’—Ê›«  «· ÿÊÌ—"
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
         TabIndex        =   27
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmExpensesInvestment.frx":76E3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   7215
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   720
      Width           =   19635
      Begin VB.TextBox TxtShareValueNew 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   7440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtShareValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   6960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtSharNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   6960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtAfterDevlopValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   6960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtCurrValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   6960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtDevlopValue 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   6960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «· ÿÊÌ—"
         ForeColor       =   &H00C00000&
         Height          =   5175
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1800
         Width           =   19215
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   4395
            Left            =   -240
            TabIndex        =   7
            Top             =   240
            Width           =   19365
            _cx             =   34158
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
            BackColorAlternate=   -2147483643
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
            Rows            =   12
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmExpensesInvestment.frx":8AE8
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
            ExplorerBar     =   7
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   6
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   4680
            Width           =   2115
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   285
            Index           =   5
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   4680
            Width           =   1515
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   3735
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   14055
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   3855
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   15615
            Begin VB.TextBox TxtCode 
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
               Left            =   11670
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox TxtRemarks 
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
               Height          =   675
               Left            =   120
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   6
               Top             =   240
               Width           =   6705
            End
            Begin VB.TextBox Text1 
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
               Left            =   11670
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   600
               Width           =   1065
            End
            Begin MSDataListLib.DataCombo DcbInvise 
               Height          =   315
               Left            =   8100
               TabIndex        =   3
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   240
               Width           =   3555
               _ExtentX        =   6271
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbLand 
               Height          =   315
               Left            =   8100
               TabIndex        =   5
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
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   0
               Left            =   6600
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   360
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «·„”«Â„…"
               Height          =   285
               Index           =   7
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·«—÷"
               Height          =   285
               Index           =   6
               Left            =   12480
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   600
               Width           =   1515
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   14055
         Begin VB.TextBox TxtFATYou 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   8400
            TabIndex        =   1
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   99155969
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmExpensesInvestment.frx":8D66
            Height          =   315
            Left            =   3600
            TabIndex        =   2
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
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
         Begin XtremeSuiteControls.RadioButton RdType 
            Height          =   255
            Index           =   0
            Left            =   2160
            TabIndex        =   77
            Top             =   240
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "„”«Â„…"
            ForeColor       =   8388608
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdType 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   78
            Top             =   240
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«—÷ „„·Êﬂ…"
            ForeColor       =   8388608
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AccountVat 
            Height          =   315
            Left            =   0
            TabIndex        =   79
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   240
            Visible         =   0   'False
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
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   6480
            TabIndex        =   48
            Top             =   240
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10050
            TabIndex        =   19
            Top             =   255
            Width           =   885
         End
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·”Â„ »⁄œ «· ÿÊÌ—"
         Height          =   285
         Index           =   9
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   7440
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·”Â„"
         Height          =   285
         Index           =   3
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   6960
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·«”Â„"
         Height          =   285
         Index           =   1
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   6960
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„… »⁄œ «· ÿÊÌ—"
         Height          =   285
         Index           =   0
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   6960
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„…«·Õ«·Ì…"
         Height          =   285
         Index           =   13
         Left            =   12960
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   6960
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «· ÿÊÌ—"
         Height          =   285
         Index           =   19
         Left            =   10200
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   6960
         Visible         =   0   'False
         Width           =   1515
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   19680
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   19680
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   21240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   19920
      TabIndex        =   30
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
      Left            =   19560
      TabIndex        =   31
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
      Height          =   1425
      Left            =   0
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7800
      Width           =   14235
      _cx             =   25109
      _cy             =   2514
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
      Begin VB.Frame Frame9 
         Height          =   690
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   0
         Width           =   4605
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   240
            Width           =   2175
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   120
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   195
            Index           =   35
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   240
            Width           =   990
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   33
         Top             =   600
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   9
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
            ButtonImage     =   "FrmExpensesInvestment.frx":8D7B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   11
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
            ButtonImage     =   "FrmExpensesInvestment.frx":F5DD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   10
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
            ButtonImage     =   "FrmExpensesInvestment.frx":F977
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   12
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
            ButtonImage     =   "FrmExpensesInvestment.frx":161D9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   13
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
            ButtonImage     =   "FrmExpensesInvestment.frx":16573
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   14
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
            ButtonImage     =   "FrmExpensesInvestment.frx":16B0D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   46
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
            ButtonImage     =   "FrmExpensesInvestment.frx":16EA7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   47
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
            ButtonImage     =   "FrmExpensesInvestment.frx":1D709
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   10320
         TabIndex        =   39
         Top             =   120
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   -840
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
      Begin ImpulseButton.ISButton ISButton3 
         Height          =   330
         Left            =   3720
         TabIndex        =   71
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
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
         ButtonImage     =   "FrmExpensesInvestment.frx":1DAA3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13200
         TabIndex        =   40
         Top             =   120
         Width           =   900
      End
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   19680
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
            Picture         =   "FrmExpensesInvestment.frx":24305
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensesInvestment.frx":2469F
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensesInvestment.frx":24A39
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensesInvestment.frx":24DD3
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensesInvestment.frx":2516D
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensesInvestment.frx":25507
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensesInvestment.frx":258A1
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmExpensesInvestment.frx":25E3B
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   19680
      TabIndex        =   41
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
      ButtonImage     =   "FrmExpensesInvestment.frx":261D5
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   20760
      TabIndex        =   44
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
      ButtonImage     =   "FrmExpensesInvestment.frx":2CA37
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   21000
      TabIndex        =   45
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
      ButtonImage     =   "FrmExpensesInvestment.frx":33299
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
      Left            =   19560
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmExpensesInvestment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
  Dim RootAccount1 As String
  Dim RootAccount2 As String
  Dim RootAccount3 As String
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecId As String
 Dim Account_Code_dynamic As String
 Dim invExpensesAcc As String
 Dim Account_Code_dynamic130 As String
 Dim CreditAcc As String
  Dim DebitAcc As String
  
 Dim II As Long
 Public LngRow  As Double
Public LngCol  As Double
Sub ClculteVAT()
If Me.TxtModFlg.Text <> "R" Then
Dim Percetage As Double
Dim account As String
PercentgValueAddedAccount_Transec XPDtbTrans.value, 44, 0, account, Percetage
TxtFATYou.Text = Percetage
AccountVat.BoundText = account
End If
End Sub
Sub GetActiveInvestInformation(Optional ID As Double = 0)
If ID <> 0 Then
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblActivateInvestment  where InviseNo=" & ID & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
TxtCurrValue.Text = IIf(IsNull(Rs4("TotalValue").value), 0, Rs4("TotalValue").value)
TxtSharNo.Text = IIf(IsNull(Rs4("SharesCount").value), 0, Rs4("SharesCount").value)
TxtShareValue.Text = IIf(IsNull(Rs4("SharesValue").value), 0, Rs4("SharesValue").value)
Else
TxtCurrValue.Text = 0
TxtShareValue.Text = 0
TxtSharNo.Text = 0
End If
End If
End Sub
Function GetActiveInvestValue(Optional ID As Double = 0, Optional invsID As Double) As Double
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
GetActiveInvestValue = 0
sql = "SELECT     SUM(DevlopValue) AS sumAfter, InvesID"
sql = sql & " From dbo.TblExpensesInvesment"
sql = sql & " Where (ID <> " & ID & ") and (InvesID = " & invsID & ")"
sql = sql & " GROUP BY InvesID"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetActiveInvestValue = IIf(IsNull(Rs4("sumAfter").value), 0, Rs4("sumAfter").value)
Else
GetActiveInvestValue = 0
End If

End Function

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200
End Sub
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords



    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
   
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    Dim SubStr2 As String
    Dim Msg1 As String
    SubStr2 = ""
       ' Msg = EleHeader.Caption & " ??? " & txtid & " EEC??I" & Date
    If TxtRemarks.Text = "" Then
        Msg1 = "„’—Ê›«  «· ÿÊÌ—»—ﬁ„ " & TxtSerial1 & "  ··„”«Â„…  " & DcbInvise.Text & "  ··«—÷ " & DcbLand.Text
 Else
 Msg1 = TxtRemarks.Text
 End If
 
        
 
 notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = 1
    
    With GridInstallments

Dim substr As String
line_no = 1
        For i = .FixedRows To .Rows - 1
    BranchID = val(dcBranch.BoundText)
 
  
  
            If val(.TextMatrix(i, .ColIndex("Valu"))) > 0 And .TextMatrix(i, .ColIndex("DevlOpID")) <> "" Then     'C?C??? C???E??E IC??
            SubStr2 = .TextMatrix(i, .ColIndex("Remarks"))
            
                substr = " ·  " & .TextMatrix(i, .ColIndex("Devlop")) & "  " & .TextMatrix(i, .ColIndex("Cus_Name")) & "   " & .TextMatrix(i, .ColIndex("Remarks"))
                If RdType(1).value = True Then
                DebitAcc = GetMyAccountCode("TblBuyLanReEst", "id", val(Me.DcbLand.BoundText), "Account_Code")
                Else
               DebitAcc = GetMyAccountCode("TblActivateInvestment", "id", val(Me.DcbInvise.BoundText), "Account_Code4")
               End If
               If .TextMatrix(i, .ColIndex("AccontCode")) = "" Then
                If .TextMatrix(i, .ColIndex("Cus_ID")) <> "" Then
                 CreditAcc = GetMyAccountCode("TblCustemers", "CusID", .TextMatrix(i, .ColIndex("Cus_ID")))
               Else
               CreditAcc = Account_Code_dynamic130
               End If
               Else
               CreditAcc = .TextMatrix(i, .ColIndex("AccontCode"))
               End If
               If Me.TxtRemarks.Text = "" Then
               Msg = Msg1 & substr
               Else
               Msg = Msg1 & SubStr2
               
               End If
                If ModAccounts.AddNewDev(LngDevID, line_no, DebitAcc, .TextMatrix(i, .ColIndex("Valu")), 0, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                   If ModAccounts.AddNewDev(LngDevID, line_no, AccountVat.BoundText, .TextMatrix(i, .ColIndex("VAT")), 0, Msg & " " & "«·ﬁÌ„… «·„÷«›…", val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                If ModAccounts.AddNewDev(LngDevID, line_no, CreditAcc, .TextMatrix(i, .ColIndex("Total")), 1, Msg, val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
  
            End If
     
     
     Next i
     
     End With
           
    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function
Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String

des = "„’—Ê›«  «· ÿÊÌ—»—ﬁ„ " & TxtSerial1 & "  ··„”«Â„…  " & DcbInvise.Text & "  ··«—÷ " & DcbLand.Text

Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblExpensesInvesment"
Filedname = "ID"
NoteSerial1 = val(TxtSerial1)
 

 notytype = 9004
Notevalue = val(lbl(6).Caption)
 

 BranchID = val(dcBranch.BoundText)
NoteDate = (XPDtbTrans.value)
 
If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TXTNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TXTNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TXTNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TXTNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TXTNoteID.Text), BranchID, user_id, NoteDate
RsSavRec.Resync adAffectCurrent
 

     End If

End Function
Private Sub DcbInvise_Change()
DcbInvise_Click (0)
End Sub

Private Sub DcbInvise_Click(Area As Integer)
Dim Dcombos As New ClsDataCombos
txtCode.Text = Me.DcbInvise.BoundText
Dcombos.GetLandActive DcbLand, val(Me.DcbInvise.BoundText)
If Me.TxtModFlg.Text <> "R" Then
If val(Me.DcbInvise.BoundText) <> 0 Then
GetActiveInvestInformation val(Me.DcbInvise.BoundText)
TxtCurrValue.Text = val(TxtCurrValue.Text) + GetActiveInvestValue(val(TxtSerial1.Text), val(Me.DcbInvise.BoundText))
End If
End If
End Sub

Private Sub DcbInvise_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 16
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End If
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
End If

End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String

    conection = "select * from TblExpensesInvesment order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
 
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetAccountingCodes AccountVat
    'Dcombos.GetLandActive DcbLand, 0
    Dcombos.GetInvestmentActive Me.DcbInvise
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
   FiLLTXT
ErrTrap:
End Sub

' save new recored or update
'++++++++++++++++++++++++++++++++++++++++
Function CheckDivide(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckDivide = False
sql = "SELECT    InvesID "
sql = sql & " From dbo.TblExpensesInvesment"
sql = sql & " Where   InvesID=" & ID & " and DivPayed=1"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckDivide = True
Else
CheckDivide = False
End If
End Function

Public Sub FiLLRec()
  
  
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.Text = "E" Then
                 StrSQL = "Delete From TblExpensesInvesmentDet Where TypTrns=1 and  ExpInvID =" & val(TxtSerial1.Text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
                     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords

   
              End If
  
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    '''////
    If RdType(1).value = True Then
RsSavRec.Fields("TypDiv").value = 1
Else
RsSavRec.Fields("TypDiv").value = 0
End If
    RsSavRec.Fields("InvesID").value = val(Me.DcbInvise.BoundText)
    RsSavRec.Fields("LandID").value = val(Me.DcbLand.BoundText)
    
 ''''//////////////////////
 
    RsSavRec.Fields("ShareValueNew").value = val(TxtShareValueNew.Text)
    RsSavRec.Fields("CurrValue").value = val(TxtCurrValue.Text)
    RsSavRec.Fields("DevlopValue").value = val(TxtDevlopValue.Text)
    RsSavRec.Fields("AfterDevlopValue").value = val(TxtAfterDevlopValue.Text)
    RsSavRec.Fields("ShareValue").value = val(Me.TxtShareValue.Text)
    RsSavRec.Fields("SharNo").value = val(Me.TxtSharNo.Text)
    RsSavRec.Fields("Remarks").value = Me.TxtRemarks.Text
    RsSavRec.Fields("Total").value = val(Me.lbl(6).Caption)
    RsSavRec.Fields("VATyo").value = val(Me.TxtFATYou.Text)
    RsSavRec.Fields("AccountCodeVat").value = AccountVat.BoundText
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ''/////
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    sql = "Update TblActivateInvestment set ExPayed=1 where InviseNo=" & val(DcbInvise.BoundText) & ""
    Cn.Execute sql
''//////////////////////////
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblExpensesInvesmentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .Rows - 1
       If val(.TextMatrix(i, .ColIndex("DevlopID"))) <> 0 Then
       RsDevsub.AddNew
                RsDevsub("TypTrns").value = 1
                RsDevsub("ExpInvID").value = val(Me.TxtSerial1.Text)
                RsDevsub("ID").value = IIf((.TextMatrix(i, .ColIndex("ExpID"))) = "", Null, val(.TextMatrix(i, .ColIndex("ExpID"))))
                RsDevsub("Cus_ID").value = IIf((.TextMatrix(i, .ColIndex("Cus_ID"))) = "", Null, val(.TextMatrix(i, .ColIndex("Cus_ID"))))
                If .TextMatrix(i, .ColIndex("AccontCode")) = "" And .TextMatrix(i, .ColIndex("Cus_ID")) <> "" Then
                .TextMatrix(i, .ColIndex("AccontCode")) = GetMyAccountCode("TblCustemers", "CusID", .TextMatrix(i, .ColIndex("Cus_ID")))
                End If
                RsDevsub("AccontCode").value = IIf((.TextMatrix(i, .ColIndex("AccontCode"))) = "", Null, (.TextMatrix(i, .ColIndex("AccontCode"))))
                RsDevsub("DevlopID").value = IIf((.TextMatrix(i, .ColIndex("DevlopID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DevlopID"))))
                RsDevsub("FromArea").value = IIf((.TextMatrix(i, .ColIndex("FromArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("FromArea"))))
                RsDevsub("ToArea").value = IIf((.TextMatrix(i, .ColIndex("ToArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("ToArea"))))
                RsDevsub("StartDate").value = IIf((.TextMatrix(i, .ColIndex("StartDate"))) = "", Null, (.TextMatrix(i, .ColIndex("StartDate"))))
                RsDevsub("EndDate").value = IIf((.TextMatrix(i, .ColIndex("EndDate"))) = "", Null, (.TextMatrix(i, .ColIndex("EndDate"))))
                RsDevsub("Valu").value = IIf((.TextMatrix(i, .ColIndex("Valu"))) = "", Null, val(.TextMatrix(i, .ColIndex("Valu"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(i, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(i, .ColIndex("Remarks"))))
                RsDevsub("Total").value = IIf((.TextMatrix(i, .ColIndex("Total"))) = "", Null, val(.TextMatrix(i, .ColIndex("Total"))))
                RsDevsub("VAT").value = IIf((.TextMatrix(i, .ColIndex("VAT"))) = "", Null, val(.TextMatrix(i, .ColIndex("VAT"))))
       RsDevsub.update
      End If
     Next i
    End With
    createVoucher
    
'''///////////////
  
      Select Case Me.TxtModFlg.Text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                
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
        Me.TXTNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)

    Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec("NoteSerial").value), "", RsSavRec("NoteSerial").value)
    Me.AccountVat.BoundText = IIf(IsNull(RsSavRec.Fields("AccountCodeVat").value), "", RsSavRec.Fields("AccountCodeVat").value)
    Me.TxtFATYou.Text = IIf(IsNull(RsSavRec.Fields("VATyo").value), 0, RsSavRec.Fields("VATyo").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbInvise.BoundText = IIf(IsNull(RsSavRec.Fields("InvesID").value), "", RsSavRec.Fields("InvesID").value)
    Me.DcbLand.BoundText = IIf(IsNull(RsSavRec.Fields("LandID").value), "", RsSavRec.Fields("LandID").value)
    TxtCurrValue.Text = IIf(IsNull(RsSavRec.Fields("CurrValue").value), 0, RsSavRec.Fields("CurrValue").value)
    TxtDevlopValue.Text = IIf(IsNull(RsSavRec.Fields("DevlopValue").value), 0, RsSavRec.Fields("DevlopValue").value)
    TxtAfterDevlopValue.Text = IIf(IsNull(RsSavRec.Fields("AfterDevlopValue").value), 0, RsSavRec.Fields("AfterDevlopValue").value)
    TxtShareValue.Text = IIf(IsNull(RsSavRec.Fields("ShareValue").value), 0, RsSavRec.Fields("ShareValue").value)
    TxtSharNo.Text = IIf(IsNull(RsSavRec.Fields("SharNo").value), 0, RsSavRec.Fields("SharNo").value)
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    lbl(6).Caption = IIf(IsNull(RsSavRec.Fields("Total").value), 0, RsSavRec.Fields("Total").value)
    TxtShareValueNew.Text = IIf(IsNull(RsSavRec.Fields("ShareValueNew").value), 0, RsSavRec.Fields("ShareValueNew").value)
    If Not (IsNull(RsSavRec.Fields("TypDiv").value)) Then
    If RsSavRec.Fields("TypDiv").value = 1 Then
    RdType(1).value = True
    Else
    RdType(0).value = True
    End If
    Else
    RdType(0).value = True
    End If
    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData
RelinGrid
ErrTrap:
End Sub
Sub maxx(Optional ByRef ExpID As Double = 0)
Dim RsDev As ADODB.Recordset
Dim StrSQL As String
Set RsDev = New ADODB.Recordset
   If ExpID <> 0 Then
     StrSQL = " select max(ExpID) as mx from FXSerialInvesment"
      RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      ExpID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FXSerialInvesment", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("ExpID").value = ExpID
RsDev.update
End If

End Sub
Function Checked(Optional ExpID As Double = 0) As Boolean
 Checked = False
  Dim RsDev As ADODB.Recordset
  Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
  If ExpID <> 0 Then
   StrSQL = " select * from FXSerialInvesment where ExpID=" & ExpID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function



Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim ExpID As Double
    With GridInstallments
        Select Case .ColKey(Col)
        


        Case "AccontName"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccontCode"), False, True)
                 .TextMatrix(Row, .ColIndex("AccontCode")) = StrAccountCode
                 
                   .TextMatrix(Row, .ColIndex("VatP")) = Me.TxtFATYou
            Case "Devlop"
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("DevlopID"), False, True)
                 .TextMatrix(Row, .ColIndex("DevlopID")) = StrAccountCode
               
   If Me.TxtModFlg.Text <> "R" Then
        If .TextMatrix(Row, .ColIndex("Devlop")) <> "" Then
        ExpID = 0
         If Me.TxtModFlg.Text = "E" Then
               ExpID = IIf(Not IsNumeric(.TextMatrix(Row, .ColIndex("ExpID"))), 0, .TextMatrix(Row, .ColIndex("ExpID")))
          
          End If
          If Me.Checked(ExpID) = True Then
        Else
       ExpID = 1
        maxx ExpID
        End If
        .TextMatrix(Row, .ColIndex("ExpID")) = ExpID
       End If

   End If
        .TextMatrix(Row, .ColIndex("VatP")) = Me.TxtFATYou
            Case "Cus_Name"
               
               
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Cus_ID"), False, True)
                 .TextMatrix(Row, .ColIndex("Cus_ID")) = StrAccountCode
                 If val(.TextMatrix(Row, .ColIndex("Cus_ID"))) <> 0 Then
                 .TextMatrix(Row, .ColIndex("AccontCode")) = GetMyAccountCode("TblCustemers", "CusID", .TextMatrix(Row, .ColIndex("Cus_ID")))
                 End If
             '''/////////
             Set rs = New ADODB.Recordset
                   StrSQL = "  select * from ACCOUNTS  where Account_Code='" & .TextMatrix(Row, .ColIndex("AccontCode")) & "'"
           
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If rs.RecordCount > 0 Then

                If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(Row, .ColIndex("AccontName")) = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                Else
                   .TextMatrix(Row, .ColIndex("AccontName")) = IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                End If
                End If
             .TextMatrix(Row, .ColIndex("VatP")) = Me.TxtFATYou
                Case "Devlop"
       End Select
   
        If Row = .Rows - 1 Then
    
          .Rows = .Rows + 1
        End If
    End With
RelinGrid
End Sub


Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With GridInstallments
        Select Case .ColKey(Col)
            Case "FromArea"
                .ComboList = ""
                 Case "ToArea"
                .ComboList = ""
                 Case "StartDate"
                .ComboList = ""
                 Case "EndDate"
                .ComboList = ""
                 Case "Valu"
                .ComboList = ""
                 Case "Remarks"
               .ComboList = ""
                 Case "VAT"
           '    Cancel = True
                 Case "Total"
               Cancel = True
        End Select
    End With
End Sub

Private Sub GridInstallments_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Me.GridInstallments
Select Case .ColKey(Col)
Case "StartDate"
 LngRow = Row
 LngCol = Col
Load FrmDateOpProject
FrmDateOpProject.Index = 17
FrmDateOpProject.show vbModal
Case "EndDate"
 LngRow = Row
 LngCol = Col
Load FrmDateOpProject
FrmDateOpProject.Index = 18
FrmDateOpProject.show vbModal
End Select
End With
End Sub

Private Sub GridInstallments_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  Account_search.show
    Account_search.case_id.Caption = 220011
 End If
End Sub

Private Sub GridInstallments_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
With GridInstallments
Select Case .ColKey(Col)
Case "AccontName"
If KeyCode = vbKeyF3 Then
  Account_search.show
    Account_search.case_id.Caption = 220011
 End If
End Select
End With
End Sub

Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    With GridInstallments

        Select Case .ColKey(Col)
   Case "StartDate"
.ColComboList(.ColIndex("StartDate")) = "..."
 Case "EndDate"
.ColComboList(.ColIndex("EndDate")) = "..."
Case "Devlop"
  StrSQL = "select * from TblInvestmentsGroup"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GridInstallments.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = GridInstallments.BuildComboList(rs, "Namee", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
Case "Cus_Name"
                   If SystemOptions.UserInterface = ArabicInterface Then
                   StrSQL = "  select CusID,CusName from TblCustemers  where type=2  order by CusName"
                   Else
                   StrSQL = "  select CusID,CusNamee from TblCustemers  where type=2  order by CusNamee"
                   End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GridInstallments.BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = GridInstallments.BuildComboList(rs, "CusNamee", "CusID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
 Case "AccontName"
                   If SystemOptions.UserInterface = ArabicInterface Then
                   StrSQL = "  select Account_Code,Account_Name from ACCOUNTS  where  last_account=1 order by Account_Name"
                   Else
                   StrSQL = "  select Account_Code,Account_NameEng from ACCOUNTS  where   last_account=1 order by Account_NameEng"
                   End If
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GridInstallments.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = GridInstallments.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

        End Select

    End With
End Sub

Private Sub ISButton3_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.Text, "170420166"
ErrTrap:
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
    Dim Total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double
    Dim account As String
  Account_Code_dynamic130 = get_account_code_branch(130, my_branch)
                                            
                If Account_Code_dynamic130 = "NO branch" Then
                    MsgBox "·„ Ì „    ÕœÌœ Õ”«» Ê”Ìÿ „’«—Ì›  ÿÊÌ—", vbCritical
                  Exit Sub
                Else

                    If Account_Code_dynamic130 = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ     Õ”«» Ê”Ìÿ „’«—Ì›  ÿÊÌ— ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                      Exit Sub
                                             
                    End If
                End If


    '---------------------- check if data Vaclete -----------------------
      If dcBranch.Text = "" And val(dcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            dcBranch.SetFocus
            Exit Sub
     End If
         '  If DcbInvise.text = "" And val(DcbInvise.BoundText) = 0 Then
         '  If SystemOptions.UserInterface = ArabicInterface Then
         '   MsgBox "⁄›Ê« ...«·—Ã«¡≈Œ Ì«— «·„”«Â„…  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
         '  Else
         '   MsgBox "Please Select Sharing ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        ' End If
        '    DcbInvise.SetFocus
        '  Exit Sub
        'End If
     If val(DcbLand.BoundText) = 0 And DcbLand.Text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·«—÷  "
     Else
     MsgBox "Please Select Land"
     End If
     DcbLand.SetFocus
     Exit Sub
     End If

        PercentgValueAddedAccount_Transec XPDtbTrans.value, 44, 0, account
      AccountVat.BoundText = account
    If AccountVat.BoundText = "" And True = True And CheckAnyVAT(XPDtbTrans.value) = True Then
    MsgBox "Ì—ÃÏ ÷»ÿ «⁄œ«œ  «·ﬁÌ„… «·„÷«›…"
    Exit Sub
    End If

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
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblExpensesInvesment", "ID", "")
    Me.TxtSerial1.Text = StrRecID
    
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
sql = " SELECT     dbo.TblExpensesInvesmentDet.ID,dbo.TblExpensesInvesmentDet.Total,dbo.TblExpensesInvesmentDet.VAT, dbo.TblExpensesInvesmentDet.ExpInvID, dbo.TblExpensesInvesmentDet.StartDate, dbo.TblExpensesInvesmentDet.EndDate, "
sql = sql + "                      dbo.TblExpensesInvesmentDet.FromArea, dbo.TblExpensesInvesmentDet.ToArea, dbo.TblExpensesInvesmentDet.Valu, dbo.TblExpensesInvesmentDet.DevlopID,"
sql = sql + "                      dbo.TblInvestmentsGroup.Code, dbo.TblInvestmentsGroup.Name, dbo.TblInvestmentsGroup.NameE, dbo.TblExpensesInvesmentDet.Remarks,"
sql = sql + "                      dbo.TblExpensesInvesmentDet.Cus_ID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
sql = sql + "                      dbo.TblExpensesInvesmentDet.AccontCode , dbo.ACCOUNTS.account_name, dbo.ACCOUNTS.account_serial, dbo.ACCOUNTS.Account_NameEng"
sql = sql + " FROM         dbo.TblExpensesInvesmentDet LEFT OUTER JOIN"
sql = sql + "                      dbo.ACCOUNTS ON dbo.TblExpensesInvesmentDet.AccontCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
sql = sql + "                      dbo.TblCustemers ON dbo.TblExpensesInvesmentDet.Cus_ID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql + "                      dbo.TblInvestmentsGroup ON dbo.TblExpensesInvesmentDet.DevlopID = dbo.TblInvestmentsGroup.ID"
sql = sql + " WHERE   dbo.TblExpensesInvesmentDet.TypTrns=1 and   (dbo.TblExpensesInvesmentDet.ExpInvID =" & val(TxtSerial1.Text) & ") "
  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("AccontCode")) = IIf(IsNull(Rs1("AccontCode").value), "", Rs1("AccontCode").value)
                   .TextMatrix(i, .ColIndex("ExpID")) = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
                   .TextMatrix(i, .ColIndex("VAT")) = IIf(IsNull(Rs1("VAT").value), 0, Rs1("VAT").value)
                   .TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(Rs1("Valu").value), "", Rs1("Valu").value)
                   .TextMatrix(i, .ColIndex("Total")) = IIf(IsNull(Rs1("Total").value), val(.TextMatrix(i, .ColIndex("Valu"))), Rs1("Total").value)
                   .TextMatrix(i, .ColIndex("StartDate")) = IIf(IsNull(Rs1("StartDate").value), Date, Rs1("StartDate").value)
                   .TextMatrix(i, .ColIndex("EndDate")) = IIf(IsNull(Rs1("EndDate").value), Date, Rs1("EndDate").value)
                   .TextMatrix(i, .ColIndex("FromArea")) = IIf(IsNull(Rs1("FromArea").value), "", Rs1("FromArea").value)
                   .TextMatrix(i, .ColIndex("ToArea")) = IIf(IsNull(Rs1("ToArea").value), "", Rs1("ToArea").value)
                   .TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(i, .ColIndex("DevlopID")) = IIf(IsNull(Rs1("DevlopID").value), "", Rs1("DevlopID").value)
                   .TextMatrix(i, .ColIndex("Cus_ID")) = IIf(IsNull(Rs1("Cus_ID").value), "", Rs1("Cus_ID").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("AccontName")) = IIf(IsNull(Rs1("Account_Name").value), "", Rs1("Account_Name").value)
                   .TextMatrix(i, .ColIndex("Cus_Name")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
                   .TextMatrix(i, .ColIndex("Devlop")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("AccontName")) = IIf(IsNull(Rs1("Account_NameEng").value), "", Rs1("Account_NameEng").value)
                   .TextMatrix(i, .ColIndex("Cus_Name")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
                   .TextMatrix(i, .ColIndex("Devlop")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub



Private Sub DcbLand_Change()
Dim Fullcode As String
If val(DcbLand.BoundText) <> 0 Then
GetTblBuyLandRealEstate val(DcbLand.BoundText), Fullcode, 0
Me.Text1.Text = Fullcode
End If
End Sub

Private Sub DcbLand_Click(Area As Integer)
DcbLand_Change
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 11
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End Sub

Private Sub lbl_Change(Index As Integer)
lbl_Click Index
End Sub

Private Sub lbl_Click(Index As Integer)
Select Case Index
Case 6
TxtDevlopValue.Text = val(lbl(6).Caption)
End Select
End Sub

Private Sub RdType_Click(Index As Integer)
Dim Dcombos As New ClsDataCombos
If RdType(0).value = True Then
   ' Dcombos.GetLandActive DcbLand
    Dcombos.GetLandActive DcbLand, 0
DcbInvise.Enabled = True
 Else
 DcbInvise.BoundText = 0
DcbInvise.Enabled = False
Dcombos.GetLandNotActive DcbLand
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim ID  As Integer
GetTblBuyLandRealEstate ID, Me.Text1.Text, 1
DcbLand.BoundText = ID
End Sub

Private Sub TxtAfterDevlopValue_Change()
If val(TxtSharNo.Text) <> 0 Then
TxtShareValueNew.Text = val(TxtAfterDevlopValue.Text) / val(TxtSharNo.Text)
TxtShareValueNew.Text = Round(TxtShareValueNew.Text, 2)
End If
End Sub



Private Sub txtCode_KeyPress(KeyAscii As Integer)
 Me.DcbInvise.BoundText = txtCode.Text
End Sub

Private Sub TxtCurrValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtAfterDevlopValue.Text = val(TxtCurrValue.Text) + val(TxtDevlopValue.Text)
TxtAfterDevlopValue.Text = Round(TxtAfterDevlopValue.Text, 2)
End If
End Sub

Private Sub TxtDevlopValue_Change()
If Me.TxtModFlg.Text <> "R" Then
TxtAfterDevlopValue.Text = val(TxtCurrValue.Text) + val(TxtDevlopValue.Text)
TxtAfterDevlopValue.Text = Round(TxtAfterDevlopValue.Text, 2)
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
    FindRec val(TxtSerial1.Text)
    Me.TxtModFlg.Text = "R"
    FiLLTXT
     BtnLast_Click
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
    Dim sql As String
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
     If CheckDivide(val(DcbInvise.BoundText)) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·Õ—ﬂ… ·«‰Â« „— »ÿ… »«· ﬁ”Ì„"
    Else
    MsgBox "Can Not Delete This process linked to with Divide the land"
    End If
    Exit Sub
    End If
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
               sql = "Update TblActivateInvestment set ExPayed=Null where InviseNo=" & val(DcbInvise.BoundText) & ""
                Cn.Execute sql
                          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
                
                
                RsSavRec.find "ID=" & val(TxtSerial1.Text), , adSearchForward, 1
                  StrSQL = "Delete From TblExpensesInvesmentDet Where TypTrns=1 and ExpInvID =" & val(TxtSerial1.Text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
                                          RsSavRec.delete
                                
                                    
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
             
               '''''''''''''''''''''''''''''''

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
               LabCurrRec.Caption = 0
     LabCountRec.Caption = 0
       ' FillGridWithData
        BtnNext_Click
     Exit Sub
ErrTrap:
     Select Case Err.Number
        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "You can not delete the record"
            StrMSG = StrMSG & " Is related to with other data"
            End If
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
    XPDtbTrans.Enabled = True
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
     XPDtbTrans.Enabled = False
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
   XPDtbTrans.Enabled = True
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
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
    If CheckDivide(val(DcbInvise.BoundText)) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·Õ—ﬂ… ·«‰Â« „— »ÿ… »«· ﬁ”Ì„"
    Else
    MsgBox "Can Not Update This process linked to with Divide the land"
    End If
    Exit Sub
    End If
        TxtModFlg = "E"
            GridInstallments.Rows = GridInstallments.Rows + 1
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
        Me.dcBranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & CHR(13)
            Msg = Msg & " You can not edit this the record now" & CHR(13)
            Msg = Msg & "It was being edited by another user on the network"
           
            End If
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
    lbl(6).Caption = 0
    TxtModFlg.Text = "N"
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 2
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = Current_branch
    dcBranch.SetFocus
  ClculteVAT
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
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
  MySQL = "SELECT        dbo.TblExpensesInvesment.ID, dbo.TblExpensesInvesment.RecordDate, dbo.TblExpensesInvesment.CurrValue, dbo.TblExpensesInvesment.DevlopValue, dbo.TblExpensesInvesment.AfterDevlopValue, "
  MySQL = MySQL & "                        dbo.TblExpensesInvesment.ShareValue, dbo.TblExpensesInvesment.SharNo, dbo.TblExpensesInvesment.Remarks, dbo.TblExpensesInvesment.Total, dbo.TblExpensesInvesment.ShareValueNew,"
  MySQL = MySQL & "                       dbo.TblExpensesInvesmentDet.StartDate, dbo.TblExpensesInvesmentDet.EndDate, dbo.TblExpensesInvesmentDet.FromArea, dbo.TblExpensesInvesmentDet.ToArea, dbo.TblExpensesInvesmentDet.Valu,"
  MySQL = MySQL & "                       dbo.TblExpensesInvesmentDet.Remarks AS RemarksDet, dbo.TblExpensesInvesment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExpensesInvesment.InvesID,"
  MySQL = MySQL & "                       dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblExpensesInvesment.LandID, dbo.TblBuyLanReEst.Name AS LandName, dbo.TblBuyLanReEst.NameE AS LandNameE,"
  MySQL = MySQL & "                       dbo.TblExpensesInvesmentDet.Cus_ID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblExpensesInvesmentDet.DevlopID,"
  MySQL = MySQL & "                       dbo.TblInvestmentsGroup.Name AS InvName, dbo.TblInvestmentsGroup.NameE AS InvNameE, dbo.TblExpensesInvesmentDet.AccontCode, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_Serial,"
  MySQL = MySQL & "                       dbo.ACCOUNTS.Account_NameEng"
  MySQL = MySQL & "   FROM            dbo.TblExpensesInvesmentDet LEFT OUTER JOIN"
  MySQL = MySQL & "                       dbo.ACCOUNTS ON dbo.TblExpensesInvesmentDet.AccontCode = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
  MySQL = MySQL & "                       dbo.TblInvestmentsGroup ON dbo.TblExpensesInvesmentDet.DevlopID = dbo.TblInvestmentsGroup.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                       dbo.TblCustemers ON dbo.TblExpensesInvesmentDet.Cus_ID = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
  MySQL = MySQL & "                       dbo.TblExpensesInvesment LEFT OUTER JOIN"
  MySQL = MySQL & "                       dbo.TblBuyLanReEst ON dbo.TblExpensesInvesment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                       dbo.Tblinvestment ON dbo.TblExpensesInvesment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
   MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblExpensesInvesment.BranchID = dbo.TblBranchesData.branch_id ON dbo.TblExpensesInvesmentDet.ExpInvID = dbo.TblExpensesInvesment.ID"
  MySQL = MySQL & " Where dbo.TblExpensesInvesmentDet.TypTrns=1 and (dbo.TblExpensesInvesment.ID =" & val(TxtSerial1.Text) & ")"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExpenseInvesment.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepExpenseInvesmentE.rpt"
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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

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


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Development  Expenses  "
    Me.lbl(4).Caption = "No"
    Me.lbl(2).Caption = "Date"
    Me.lbl(7).Caption = "Branch"
    Label1(7).Caption = "Investment"
    Me.Label1(2).Caption = Me.Caption
    Label1(0).Caption = "Remarks"
    Label1(6).Caption = "Land"
    lbl(5).Caption = "Total"
    lbl(13).Caption = "Curr. Value"
    lbl(19).Caption = "Develop. Value"
    lbl(0).Caption = "After Develop"
    lbl(1).Caption = "No.Share"
    lbl(3).Caption = "Share Value"
    lbl(9).Caption = "Share New Value"
    Frame6.Caption = "Data of Development"
    ISButton5.Caption = "Print"
    ISButton8.Caption = "Search"
    '''''''''''''' next
    ''''''''''''''''''''''''''''''''''''''' next
    Me.Label2(0).Caption = "Current Record"
    Me.Label2(1).Caption = "No. Recordes"
    Me.lbl(8).Caption = "by"
    '''''''''''''''''''''''''''''''' next
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    BtnUpdate.Caption = "Refresh "
    ISButton1.Caption = "Print"
    btnQuery.Caption = "Search"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    
  With Me.GridInstallments
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("Devlop")) = "Type Development"
  .TextMatrix(0, .ColIndex("FromArea")) = "From Area"
  .TextMatrix(0, .ColIndex("ToArea")) = "To Area"
  .TextMatrix(0, .ColIndex("Cus_Name")) = "Developer"
  .TextMatrix(0, .ColIndex("StartDate")) = "Start Date"
  .TextMatrix(0, .ColIndex("EndDate")) = "End Date"
  .TextMatrix(0, .ColIndex("Valu")) = "Value"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
  End With
ErrTrap:
End Sub

Sub RelinGrid()
Dim Sm, summation As Double
Dim Counter As Integer
Dim i As Integer
Counter = 0
Sm = 0
summation = 0
lbl(6).Caption = 0
With Me.GridInstallments
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("Valu"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
summation = summation + val(.TextMatrix(i, .ColIndex("Valu")))
   If .TextMatrix(i, .ColIndex("VatP")) = "" Then .TextMatrix(i, .ColIndex("VatP")) = Me.TxtFATYou
        If val(TxtFATYou.Text) > 0 Then
        .TextMatrix(i, .ColIndex("VAT")) = val(.TextMatrix(i, .ColIndex("Valu"))) * val(.TextMatrix(i, .ColIndex("VatP"))) / 100
        Else
        .TextMatrix(i, .ColIndex("VAT")) = 0
        End If
        .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("VAT"))) + val(.TextMatrix(i, .ColIndex("Valu")))
End If
Next i
lbl(6).Caption = summation

End With
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblExpensesInvesment"
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
Private Sub TxtSharNo_Change()
If val(TxtSharNo.Text) <> 0 Then
TxtShareValueNew.Text = Round(val(TxtAfterDevlopValue.Text) / val(TxtSharNo.Text), 2)
End If
End Sub

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg <> "R" Then
TxtNoteSerial.Text = ""
ClculteVAT
End If


End Sub
