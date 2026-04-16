VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmDiviInvestment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmDiviInvestment.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8535
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
      TabIndex        =   34
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmDiviInvestment.frx":6852
      Left            =   15480
      List            =   "FrmDiviInvestment.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   33
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
      TabIndex        =   27
      Top             =   0
      Width           =   14505
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox tXTRootAccount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   28
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
         ButtonImage     =   "FrmDiviInvestment.frx":687B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   29
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
         ButtonImage     =   "FrmDiviInvestment.frx":6C15
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   30
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
         ButtonImage     =   "FrmDiviInvestment.frx":6FAF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   31
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
         ButtonImage     =   "FrmDiviInvestment.frx":7349
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ﬁ”Ì„ «·«—«÷Ì"
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
         TabIndex        =   32
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmDiviInvestment.frx":76E3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   720
      Width           =   14235
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00C00000&
         Height          =   4215
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   2280
         Width           =   14055
         Begin VB.TextBox TxtSharMetre 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   3720
            Width           =   1215
         End
         Begin VB.TextBox TxtMetreValue 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   4080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtAlwAreaAfter 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox TxtAlwArea 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   3720
            Width           =   1215
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   2955
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   13845
            _cx             =   24421
            _cy             =   5212
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
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmDiviInvestment.frx":8AE8
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
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   330
            Left            =   12600
            TabIndex        =   88
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   3240
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› ’›"
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
            ButtonImage     =   "FrmDiviInvestment.frx":8D61
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   330
            Left            =   10920
            TabIndex        =   89
            ToolTipText     =   "Õ–› «·ﬂ·"
            Top             =   3240
            Width           =   1335
            _ExtentX        =   2355
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
            ButtonImage     =   "FrmDiviInvestment.frx":F5C3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ —"
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   14
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   3720
            Width           =   315
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Â„ Ì”«ÊÌ"
            Height          =   285
            Index           =   12
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   3720
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·„ —"
            Height          =   285
            Index           =   11
            Left            =   3360
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   3720
            Visible         =   0   'False
            Width           =   1875
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”«Õ… «·„ «Õ… »⁄œ «· ﬁ”Ì„"
            Height          =   285
            Index           =   10
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   3720
            Width           =   1995
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”«Õ… «·„ «Õ… ·· ﬁ”Ì„"
            Height          =   285
            Index           =   9
            Left            =   12000
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   3720
            Width           =   1755
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   285
            Index           =   6
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   3720
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·„”«Õ… «·„” ﬁÿ⁄…"
            Height          =   285
            Index           =   5
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   3720
            Width           =   1875
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   3735
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   14055
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   3855
            Left            =   0
            TabIndex        =   54
            Top             =   0
            Width           =   14055
            Begin VB.TextBox TxtCodeUnit 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtDivArae 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   600
               Width           =   1215
            End
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
               TabIndex        =   74
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox TxtDevlopValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtCurrValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   5340
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox TxtAfterDevlopValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtSharNo 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtShareValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtRemarks 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Height          =   675
               Left            =   60
               MaxLength       =   50
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   960
               Width           =   12675
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
               TabIndex        =   7
               Top             =   600
               Width           =   1065
            End
            Begin MSDataListLib.DataCombo DcbInvise 
               Height          =   315
               Left            =   8160
               TabIndex        =   3
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   240
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbLand 
               Height          =   315
               Left            =   8160
               TabIndex        =   8
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   600
               Width           =   3435
               _ExtentX        =   6059
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDivMain 
               Height          =   315
               Left            =   4440
               TabIndex        =   80
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   600
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·ÊÕœ…"
               Height          =   285
               Index           =   18
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„”«Õ…"
               Height          =   285
               Index           =   17
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÊÕœ… «·—∆Ì”Ì…"
               Height          =   285
               Index           =   16
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   600
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   15
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   240
               Width           =   5235
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «· ÿÊÌ—"
               Height          =   285
               Index           =   19
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   240
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…«·Õ«·Ì…"
               Height          =   285
               Index           =   13
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„… »⁄œ «· ÿÊÌ—"
               Height          =   285
               Index           =   0
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   240
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·«”Â„"
               Height          =   285
               Index           =   1
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   600
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁÌ„… «·”Â„"
               Height          =   285
               Index           =   3
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   600
               Visible         =   0   'False
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ« "
               Height          =   285
               Index           =   0
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   1200
               Width           =   1515
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " «·„”«Â„…"
               Height          =   285
               Index           =   7
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   61
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
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   600
               Width           =   1515
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   14055
         Begin XtremeSuiteControls.RadioButton RdType 
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   86
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
            Left            =   8760
            TabIndex        =   1
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   208535553
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmDiviInvestment.frx":15E25
            Height          =   315
            Left            =   3960
            TabIndex        =   2
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
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
            Index           =   1
            Left            =   360
            TabIndex        =   87
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   7560
            TabIndex        =   53
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
            TabIndex        =   25
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
            TabIndex        =   24
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
      RightToLeft     =   -1  'True
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   35
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
      TabIndex        =   36
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
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   7200
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
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   38
         Top             =   600
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   14
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
            ButtonImage     =   "FrmDiviInvestment.frx":15E3A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   16
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
            ButtonImage     =   "FrmDiviInvestment.frx":1C69C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   15
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
            ButtonImage     =   "FrmDiviInvestment.frx":1CA36
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   17
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
            ButtonImage     =   "FrmDiviInvestment.frx":23298
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   18
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
            ButtonImage     =   "FrmDiviInvestment.frx":23632
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   19
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
            ButtonImage     =   "FrmDiviInvestment.frx":23BCC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   51
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
            ButtonImage     =   "FrmDiviInvestment.frx":23F66
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   52
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
            ButtonImage     =   "FrmDiviInvestment.frx":2A7C8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9840
         TabIndex        =   44
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
         TabIndex        =   48
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
         Left            =   3840
         TabIndex        =   78
         ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
         Top             =   120
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
         ButtonImage     =   "FrmDiviInvestment.frx":2AB62
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13080
         TabIndex        =   45
         Top             =   120
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
            Picture         =   "FrmDiviInvestment.frx":313C4
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiviInvestment.frx":3175E
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiviInvestment.frx":31AF8
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiviInvestment.frx":31E92
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiviInvestment.frx":3222C
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiviInvestment.frx":325C6
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiviInvestment.frx":32960
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDiviInvestment.frx":32EFA
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   46
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
      ButtonImage     =   "FrmDiviInvestment.frx":33294
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   49
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
      ButtonImage     =   "FrmDiviInvestment.frx":39AF6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   50
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
      ButtonImage     =   "FrmDiviInvestment.frx":40358
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
      TabIndex        =   47
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmDiviInvestment"
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
 Dim ii As Long
 Public LonRow As Double
Public LngCol As Double


Private Sub DcbInvise_Change()
DcbInvise_Click (0)
End Sub

Private Sub DcbInvise_Click(Area As Integer)
Dim InvestTotal As Double
Dim CountShare As Double
Dim Dcombos As New ClsDataCombos
TxtCode.text = Me.DcbInvise.BoundText
Dcombos.GetLandActive DcbLand, val(Me.DcbInvise.BoundText)
If Me.TxtModFlg.text <> "R" Then
If val(DcbInvise.BoundText) <> 0 Then
GetInvestInformation val(Me.DcbInvise.BoundText), InvestTotal, CountShare
TxtCurrValue.text = InvestTotal
TxtSharNo.text = CountShare
TxtAlwArea.text = GetTOtalArea()
TxtAlwArea.text = val(TxtAlwArea.text) - SumValuTotal(val(TxtSerial1.text), val(DcbInvise.BoundText))

End If
End If
End Sub
Function GetTOtalArea() As Double
Dim Rs5 As ADODB.Recordset
Dim sql As String
Set Rs5 = New ADODB.Recordset
sql = sql & " SELECT     SUM(Area) AS SumTotalArea, InviseOrder"
sql = sql & " From dbo.TblActivateInvestment"
sql = sql & " Where (InviseOrder = " & val(DcbInvise.BoundText) & ") "
sql = sql & " GROUP BY InviseOrder"
Rs5.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
GetTOtalArea = IIf(IsNull(Rs5("SumTotalArea").value), 0, Rs5("SumTotalArea").value)
Else
GetTOtalArea = 0
End If
End Function
Function CheckSalesdET(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckSalesdET = False
sql = "SELECT    InvID "
sql = sql & " From dbo.TblDivInvestInformation "
sql = sql & " Where   DivIDDet=" & ID & " and (SalesPayed=1 or SalesBlocPayed=1) "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckSalesdET = True
Else
CheckSalesdET = False
End If
End Function
Function CheckSales(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckSales = False
sql = "SELECT    InvID "
sql = sql & " From dbo.TblDivInvestInformation "
sql = sql & " Where   InvID=" & ID & " and (SalesPayed=1 or SalesBlocPayed=1) "
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckSales = True
Else
CheckSales = False
End If
End Function
Function CheckBuy(Optional ID As Double = 0) As Boolean
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
CheckBuy = False
sql = "SELECT    InvesID "
sql = sql & " From dbo.TblDivInvesment"
sql = sql & " Where   InvesID=" & ID & " and BuyPayed=1"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
CheckBuy = True
Else
CheckBuy = False
End If
End Function

Private Sub DcbInvise_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearchinvestment.inde = 17
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End If
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
   If SystemOptions.UserInterface = ArabicInterface Then
                GridInstallments.ColComboList(GridInstallments.ColIndex("EffectID")) = "#1; „ «Õ ··»Ì⁄|#2;  €Ì— „ «Õ"
                   GridInstallments.ColComboList(GridInstallments.ColIndex("HaveDetail")) = "#1;  »œÊ‰ |#2;  ·Â "
          
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               GridInstallments.ColComboList(GridInstallments.ColIndex("EffectID")) = "#1;Effect |#2;No Effect "
               GridInstallments.ColComboList(GridInstallments.ColIndex("HaveDetail")) = "#1;Without  |#2;Have  "
            End If
            
    conection = "select * from TblDivInvesment order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me
 
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetTblSpreadingInvestment DcbDivMain, 1
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetLandActive DcbLand
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
Public Sub FiLLRec()
  '  On Error GoTo ErrTrap
    Dim sql As String
    Dim ID As Double
             If Me.TxtModFlg.text = "E" Then
                 StrSQL = "Delete From TblDivInvesmentDet Where DivInvID =" & val(TxtSerial1.text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If
  
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("BranchID").value = val(Me.dcBranch.BoundText)
    '''////
    RsSavRec.Fields("SharMetre").value = val(TxtSharMetre.text)
    RsSavRec.Fields("InvesID").value = val(Me.DcbInvise.BoundText)
    RsSavRec.Fields("LandID").value = val(Me.DcbLand.BoundText)
    RsSavRec.Fields("MeterValue").value = val(Me.TxtMetreValue.text)
    RsSavRec.Fields("SharMetre").value = val(TxtSharMetre.text)
 ''''//////////////////////
    RsSavRec.Fields("CurrValue").value = val(TxtCurrValue.text)
    RsSavRec.Fields("DevlopValue").value = val(TxtDevlopValue.text)
    RsSavRec.Fields("AfterDevlopValue").value = val(TxtAfterDevlopValue.text)
    RsSavRec.Fields("ShareValue").value = val(Me.TxtShareValue.text)
    RsSavRec.Fields("SharNo").value = val(Me.TxtSharNo.text)
    RsSavRec.Fields("Remarks").value = Me.TxtRemarks.text
    RsSavRec.Fields("Total").value = val(Me.lbl(6).Caption)
    RsSavRec.Fields("AlwArea").value = val(TxtAlwArea.text)
    RsSavRec.Fields("AlwAreaAfter").value = val(TxtAlwAreaAfter.text)
    RsSavRec.Fields("DivMainID").value = val(Me.DcbDivMain.BoundText)
    RsSavRec.Fields("DivArae").value = val(Me.TxtDivArae.text)
    RsSavRec.Fields("CodeUnit").value = (Me.TxtCodeUnit.text)
If RdType(1).value = True Then
RsSavRec.Fields("TypDiv").value = 1
Else
RsSavRec.Fields("TypDiv").value = 0
End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ''/////
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update
    sql = "Update TblExpensesInvesment set DivPayed=1 where InvesID=" & val(DcbInvise.BoundText) & " "
    Cn.Execute sql
''//////////////////////////
Dim HaveDetail As Integer
Dim EffectID As Integer
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblDivInvesmentDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim i As Integer
    Dim str2 As String
    With Me.GridInstallments
       For i = .FixedRows To .rows - 1
       If val(.TextMatrix(i, .ColIndex("TypeDivi"))) <> 0 Then
       RsDevsub.AddNew
       
                RsDevsub("DivInvID").value = val(TxtSerial1.text)
                RsDevsub("TypBP").value = IIf((.TextMatrix(i, .ColIndex("TypBP"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypBP"))))
                RsDevsub("ID").value = IIf((.TextMatrix(i, .ColIndex("DiviID"))) = "", Null, val(.TextMatrix(i, .ColIndex("DiviID"))))
                RsDevsub("HaveDetail").value = IIf((.TextMatrix(i, .ColIndex("HaveDetail"))) = "", Null, val(.TextMatrix(i, .ColIndex("HaveDetail"))))
                RsDevsub("TypeDivi").value = IIf((.TextMatrix(i, .ColIndex("TypeDivi"))) = "", Null, val(.TextMatrix(i, .ColIndex("TypeDivi"))))
                RsDevsub("EffectID").value = IIf((.TextMatrix(i, .ColIndex("EffectID"))) = "", Null, val(.TextMatrix(i, .ColIndex("EffectID"))))
                RsDevsub("Area").value = IIf((.TextMatrix(i, .ColIndex("Area"))) = "", Null, val(.TextMatrix(i, .ColIndex("Area"))))
                RsDevsub("TotalArea").value = IIf((.TextMatrix(i, .ColIndex("TotalArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("TotalArea"))))
                RsDevsub("NewArea").value = IIf((.TextMatrix(i, .ColIndex("NewArea"))) = "", Null, val(.TextMatrix(i, .ColIndex("NewArea"))))
                RsDevsub("PartNo").value = IIf((.TextMatrix(i, .ColIndex("PartNo"))) = "", Null, (.TextMatrix(i, .ColIndex("PartNo"))))
                RsDevsub("BlokNo").value = IIf((.TextMatrix(i, .ColIndex("BlokNo"))) = "", Null, (.TextMatrix(i, .ColIndex("BlokNo"))))
                RsDevsub("Nourth").value = IIf((.TextMatrix(i, .ColIndex("Nourth"))) = "", Null, (.TextMatrix(i, .ColIndex("Nourth"))))
                RsDevsub("South").value = IIf((.TextMatrix(i, .ColIndex("South"))) = "", Null, (.TextMatrix(i, .ColIndex("South"))))
                RsDevsub("East").value = IIf((.TextMatrix(i, .ColIndex("East"))) = "", Null, (.TextMatrix(i, .ColIndex("East"))))
                RsDevsub("West").value = IIf((.TextMatrix(i, .ColIndex("West"))) = "", Null, (.TextMatrix(i, .ColIndex("West"))))
                RsDevsub("StraInform").value = IIf((.TextMatrix(i, .ColIndex("StraInform"))) = "", Null, (.TextMatrix(i, .ColIndex("StraInform"))))
       RsDevsub.update
        If val(.TextMatrix(i, .ColIndex("EffectID"))) = 2 Then
        EffectID = 0
        Else
        EffectID = 1
       End If
         If val(.TextMatrix(i, .ColIndex("HaveDetail"))) = 2 Then
        HaveDetail = 1
        Else
        HaveDetail = 0
        End If
       If val(.TextMatrix(i, .ColIndex("HaveDetail"))) = 2 And val(.TextMatrix(i, .ColIndex("EffectID"))) = 1 Then
       saveDetails i, RsDevsub("id").value, IIf((.TextMatrix(i, .ColIndex("PartNo"))) = "", Null, (.TextMatrix(i, .ColIndex("PartNo")))), val(.TextMatrix(i, .ColIndex("TypeDivi")))
       Else
       saveDetails2 RsDevsub("id").value, (.TextMatrix(i, .ColIndex("Nourth"))), HaveDetail, EffectID, (.TextMatrix(i, .ColIndex("South"))), (.TextMatrix(i, .ColIndex("East"))), (.TextMatrix(i, .ColIndex("West"))), val((.TextMatrix(i, .ColIndex("Area")))), val((.TextMatrix(i, .ColIndex("TotalArea")))), val((.TextMatrix(i, .ColIndex("NewArea")))), ((.TextMatrix(i, .ColIndex("PartNo")))), val(.TextMatrix(i, .ColIndex("TypeDivi")))
       End If
      End If
     Next i
    End With
'''///////////////
  
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & CHR(13)
                Msg = Msg + " You want to enter another record?"
           End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                 If SystemOptions.UserInterface = ArabicInterface Then
             Else
              
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
                MsgBox "Changes Was Saved ... Continuation Add Data ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            End If
                Call btnNew_Click
            Else
                Me.Refresh
                TxtModFlg = "R"
                FiLLTXT
            End If
         Case "E"
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
                Me.Refresh
                FiLLTXT
                TxtModFlg = "R"
            Else
                MsgBox "Changes was saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                
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
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbInvise.BoundText = IIf(IsNull(RsSavRec.Fields("InvesID").value), "", RsSavRec.Fields("InvesID").value)
    Me.DcbLand.BoundText = IIf(IsNull(RsSavRec.Fields("LandID").value), "", RsSavRec.Fields("LandID").value)
    TxtCurrValue.text = IIf(IsNull(RsSavRec.Fields("CurrValue").value), 0, RsSavRec.Fields("CurrValue").value)
    TxtDevlopValue.text = IIf(IsNull(RsSavRec.Fields("DevlopValue").value), 0, RsSavRec.Fields("DevlopValue").value)
    TxtAfterDevlopValue.text = IIf(IsNull(RsSavRec.Fields("AfterDevlopValue").value), 0, RsSavRec.Fields("AfterDevlopValue").value)
    TxtShareValue.text = IIf(IsNull(RsSavRec.Fields("ShareValue").value), 0, RsSavRec.Fields("ShareValue").value)
    TxtSharNo.text = IIf(IsNull(RsSavRec.Fields("SharNo").value), 0, RsSavRec.Fields("SharNo").value)
    TxtRemarks.text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    lbl(6).Caption = IIf(IsNull(RsSavRec.Fields("Total").value), 0, RsSavRec.Fields("Total").value)
    TxtAlwArea.text = IIf(IsNull(RsSavRec.Fields("AlwArea").value), 0, RsSavRec.Fields("AlwArea").value)
    TxtAlwAreaAfter.text = IIf(IsNull(RsSavRec.Fields("AlwAreaAfter").value), 0, RsSavRec.Fields("AlwAreaAfter").value)
    Me.TxtMetreValue.text = IIf(IsNull(RsSavRec.Fields("MeterValue").value), 0, RsSavRec.Fields("MeterValue").value)
    Me.TxtSharMetre.text = IIf(IsNull(RsSavRec.Fields("SharMetre").value), 0, RsSavRec.Fields("SharMetre").value)
    Me.TxtSharMetre.text = IIf(IsNull(RsSavRec.Fields("SharMetre").value), 0, RsSavRec.Fields("SharMetre").value)
    Me.DcbDivMain.BoundText = IIf(IsNull(RsSavRec.Fields("DivMainID").value), 0, RsSavRec.Fields("DivMainID").value)
    Me.TxtDivArae.text = IIf(IsNull(RsSavRec.Fields("DivArae").value), 0, RsSavRec.Fields("DivArae").value)
    TxtCodeUnit.text = IIf(IsNull(RsSavRec.Fields("CodeUnit").value), "", RsSavRec.Fields("CodeUnit").value)
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
Sub maxx(Optional ByRef DivID As Double = 0)
Dim RsDev As ADODB.Recordset
Dim StrSQL As String
Set RsDev = New ADODB.Recordset
   If DivID <> 0 Then
     StrSQL = " select max(DivID) as mx from FXSerialInvesment"
      RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      DivID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FXSerialInvesment", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("DivID").value = DivID
RsDev.update
End If

End Sub
Function Checked(Optional DivID As Double = 0) As Boolean
 Checked = False
  Dim RsDev As ADODB.Recordset
  Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
  If DivID <> 0 Then
   StrSQL = " select * from FXSerialInvesment where DivID=" & DivID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function

Public Sub GridInstallments_AfterEdit(ByVal row As Long, ByVal Col As Long)
Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim DiviID As Double
    With GridInstallments
        Select Case .ColKey(Col)
            Case "EffectID"
               If row = 1 Then
                 .TextMatrix(row, .ColIndex("TotalArea")) = val(TxtDivArae.text)
                 Else
                 .TextMatrix(row, .ColIndex("TotalArea")) = val(.TextMatrix(row - 1, .ColIndex("NewArea")))
                 End If
        If val(.TextMatrix(row, .ColIndex("EffectID"))) = 2 Then
        If val(.TextMatrix(row, .ColIndex("TotalArea"))) >= val(.TextMatrix(row, .ColIndex("Area"))) Then
        .TextMatrix(row, .ColIndex("NewArea")) = val(.TextMatrix(row, .ColIndex("TotalArea"))) - val(.TextMatrix(row, .ColIndex("Area")))
        Else
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·„”«Õ… «ﬂ»— „‰ «·„”«Õ… «·«Ã„«·Ì…"
        Else
        MsgBox "Can not Area larger than total Area"
        End If
        .TextMatrix(row, .ColIndex("Area")) = 0
        Exit Sub
        End If
        Else
        .TextMatrix(row, .ColIndex("NewArea")) = val(.TextMatrix(row, .ColIndex("TotalArea")))
        End If
        
        Case "Area"
'        If val(.TextMatrix(row, .ColIndex("Area"))) > val(.TextMatrix(row, .ColIndex("TotalArea"))) Then
'          If SystemOptions.UserInterface = ArabicInterface Then
'        MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·„”«Õ… «ﬂ»— „‰ «·„”«Õ… «·«Ã„«·Ì…"
'        Else
'        MsgBox "Can not Area larger than total Area"
'        End If
'        .TextMatrix(row, .ColIndex("Area")) = 0
'        Exit Sub
        
     '   End If
      If val(.TextMatrix(row, .ColIndex("EffectID"))) = 2 Then
        If val(.TextMatrix(row, .ColIndex("TotalArea"))) >= val(.TextMatrix(row, .ColIndex("Area"))) Then
        .TextMatrix(row, .ColIndex("NewArea")) = val(.TextMatrix(row, .ColIndex("TotalArea"))) - val(.TextMatrix(row, .ColIndex("Area")))
        Else
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·„”«Õ… «ﬂ»— „‰ «·„”«Õ… «·«Ã„«·Ì…"
        Else
        MsgBox "Can not Area larger than total Area"
        End If
        .TextMatrix(row, .ColIndex("Area")) = 0
        Exit Sub
        End If
        Else
        .TextMatrix(row, .ColIndex("NewArea")) = val(.TextMatrix(row, .ColIndex("TotalArea")))
        End If
     
            Case "Name"
             
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TypeDivi"), False, True)
                 .TextMatrix(row, .ColIndex("TypeDivi")) = StrAccountCode
                 If row = 1 Then
                 .TextMatrix(row, .ColIndex("TotalArea")) = val(TxtAlwArea.text)
                 Else
                 .TextMatrix(row, .ColIndex("TotalArea")) = val(.TextMatrix(row - 1, .ColIndex("NewArea")))
                 End If
                      If Me.TxtModFlg.text <> "R" Then
        If .TextMatrix(row, .ColIndex("TypeDivi")) <> "" Then
        DiviID = 0
         If Me.TxtModFlg.text = "E" Then
               DiviID = IIf(Not IsNumeric(.TextMatrix(row, .ColIndex("DiviID"))), 0, .TextMatrix(row, .ColIndex("DiviID")))
          
          End If
          If Me.Checked(DiviID) = True Then
        Else
       DiviID = 1
        maxx DiviID
        End If
        .TextMatrix(row, .ColIndex("DiviID")) = DiviID
       End If

 End If
           
           End Select
   
        If row = .rows - 1 Then
    
          .rows = .rows + 1
        End If
    End With
RelinGrid
End Sub
Sub saveDetails2(Optional DivIDDet As Double = 0, Optional Nourth As String, Optional TypeTr As Integer, Optional EffectID As Integer _
, Optional South As String, Optional East As String, Optional West As String, Optional Area As Double, Optional TotalArea As Double _
, Optional NewArea As Double, Optional PartNo As String, Optional TypeDivi As Double = 0)
Dim RsDetails11 As ADODB.Recordset
 Dim IDDet As Double
Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
Dim st As String
Dim nElements As Integer
Dim k, m As Integer
Dim Diff As Integer
If DivIDDet <> 0 Then
Set RsDetails11 = New ADODB.Recordset
If Me.TxtModFlg.text = "E" Then
StrSQL = "delete From TblDivInvestInformation  where  DivIDDet =" & DivIDDet
                   Cn.Execute StrSQL, , adExecuteNoRecords
End If
    StrSQL = "SELECT  *  from dbo.TblDivInvestInformation Where (1 = -1)"
   RsDetails11.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        RsDetails11.AddNew
         RsDetails11("DivID").value = val(TxtSerial1.text)
         If Me.RdType(1).value = True Then
         RsDetails11("TypDiv").value = 1
         RsDetails11("InvID").value = val(DcbLand.BoundText)
         Else
         RsDetails11("TypDiv").value = 0
         RsDetails11("InvID").value = val(DcbInvise.BoundText)
         End If
         RsDetails11("DivIDDet").value = DivIDDet
         RsDetails11("RecordDate").value = XPDtbTrans.value
         RsDetails11("TypeTr").value = TypeTr
         RsDetails11("EffectID").value = EffectID
         RsDetails11("TypeDivi").value = TypeDivi
         RsDetails11("PartNo").value = PartNo
         RsDetails11("BlokNo").value = PartNo
         RsDetails11("Nourth").value = Nourth
         RsDetails11("South").value = South
         RsDetails11("East").value = East
         RsDetails11("West").value = West
         RsDetails11("Area").value = Area
         RsDetails11("TotalArea").value = TotalArea
         RsDetails11("NewArea").value = NewArea
         RsDetails11("DivMainID").value = val(DcbDivMain.BoundText)
         RsDetails11("CodeUnit").value = TxtCodeUnit.text
         RsDetails11.update
End If

End Sub
Sub saveDetails(Optional i As Integer = 0, Optional DivIDDet As Double = 0, Optional PartNo As String, Optional TypeDivi As Double = 0)

Dim RsDetails11 As ADODB.Recordset
 Dim IDDet As Double
Dim astrSplit2tems2() As String
Dim astrSplitItems() As String
Dim j As Integer
Dim st As String
Dim nElements As Integer
Dim k, m As Integer
Dim Diff As Integer
Dim RsDetails12 As ADODB.Recordset
Set RsDetails12 = New ADODB.Recordset
Dim IDDet1 As Double
If DivIDDet <> 0 Then
Set RsDetails11 = New ADODB.Recordset
If Me.TxtModFlg.text = "E" Then
StrSQL = "delete From TblDivInvesmentDetCheld  where  DivIDDet =" & DivIDDet
Cn.Execute StrSQL, , adExecuteNoRecords
StrSQL = "delete From TblDivInvestInformation  where  DivIDDet =" & DivIDDet
                   Cn.Execute StrSQL, , adExecuteNoRecords
End If
    StrSQL = "SELECT  *  from dbo.TblDivInvesmentDetCheld Where (1 = -1)"
   RsDetails11.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   StrSQL = "SELECT  *  from dbo.TblDivInvestInformation Where (1 = -1)"
   RsDetails12.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

k = 0
     If GridInstallments.TextMatrix(i, GridInstallments.ColIndex("StraInform")) <> "" Then
          st = GridInstallments.TextMatrix(i, GridInstallments.ColIndex("StraInform"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
   
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
        
         For j = 0 To nElements - 1
          RsDetails11.AddNew
          RsDetails12.AddNew
         astrSplit2tems2 = Split(astrSplitItems(j), "#")
         Diff = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
         Diff = Diff / 8
         m = 0
         
        For k = 0 To Diff - 1
        '''''
        
         RsDetails12("DivID").value = val(TxtSerial1.text)
         RsDetails12("DivIDDet").value = DivIDDet
         RsDetails12("RecordDate").value = XPDtbTrans.value
         RsDetails12("TypeTr").value = 1
         RsDetails12("EffectID").value = 1
         RsDetails12("PartNo").value = PartNo
         If Me.RdType(1).value = True Then
         RsDetails12("InvID").value = val(DcbLand.BoundText)
         Else
         RsDetails12("InvID").value = val(DcbInvise.BoundText)
         End If
         
         RsDetails12("DivMainID").value = val(DcbDivMain.BoundText)
         RsDetails12("CodeUnit").value = TxtCodeUnit.text
        Dim Replac As String
        ''''
        RsDetails12("TypeDivi").value = TypeDivi
        RsDetails11("TypeDivi").value = TypeDivi
        
        RsDetails11("DivID").value = val(TxtSerial1.text)
         RsDetails11("DivIDDet").value = DivIDDet
        ' RsDetails11("RecDate").value = astrSplit2tems2(m)
        ' m = m + 1
        
         Replac = Replace(Replace(astrSplit2tems2(m), CHR(10), ""), CHR(13), "")
         Replac = Trim(Replac)
         RsDetails11("BlokNo").value = Replac
         RsDetails12("BlokNo").value = Replac
         m = m + 1
         RsDetails11("unitunid").value = val(astrSplit2tems2(m))
         RsDetails12("unitunid").value = val((astrSplit2tems2(m)))
         m = m + 1
         
         RsDetails11("Nourth").value = astrSplit2tems2(m)
         RsDetails12("Nourth").value = (astrSplit2tems2(m))
         m = m + 1
         RsDetails11("South").value = (astrSplit2tems2(m))
         RsDetails12("South").value = (astrSplit2tems2(m))
         m = m + 1
         RsDetails11("East").value = (astrSplit2tems2(m))
         RsDetails12("East").value = (astrSplit2tems2(m))
         m = m + 1
         RsDetails11("West").value = (astrSplit2tems2(m))
         RsDetails12("West").value = (astrSplit2tems2(m))
         m = m + 1
         RsDetails11("Area").value = val((astrSplit2tems2(m)))
         RsDetails12("Area").value = val(astrSplit2tems2(m))
         m = m + 1
         RsDetails11("TotalArea").value = val((astrSplit2tems2(m)))
         RsDetails12("TotalArea").value = val(astrSplit2tems2(m))
         m = m + 1
         RsDetails11("NewArea").value = val((astrSplit2tems2(m)))
         RsDetails12("NewArea").value = val(astrSplit2tems2(m))
         m = m + 1
         RsDetails11.update
         IDDet1 = RsDetails11("ID").value
          RsDetails12("DivIDDet2").value = IDDet1
         RsDetails12.update
      Next k
       Next j
          End If
End If

End Sub

Private Sub GridInstallments_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
 With GridInstallments
        Select Case .ColKey(Col)
            Case "Nourth"
                .ComboList = ""
                 Case "South"
                .ComboList = ""
                 Case "East"
                .ComboList = ""
                 Case "West"
                .ComboList = ""
                 
                   Case "Area"
                   
           If val(.TextMatrix(row, .ColIndex("EffectID"))) = 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «· «ÀÌ— «Ê·«"
           Else
           MsgBox "Please seletc Tpe of Effect"
           End If
           Cancel = True
           Exit Sub
           Else
           Cancel = False
           .ComboList = ""
           End If
                
                 Case "TotalArea"
               Cancel = True
                 Case "NewArea"
               Cancel = True
                 Case "BlokNo"
               .ComboList = ""
                 Case "PartNo"
               .ComboList = ""
        End Select
    End With
End Sub

Private Sub GridInstallments_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With GridInstallments
Select Case .ColKey(Col)
Case "BlokNo"

 LonRow = row
 If val(.TextMatrix(row, .ColIndex("Area"))) = 0 Then
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "Ì—ÃÏ «œŒ«· «·„”«Õ… «·„—«œ  ﬁ”Ì„Â« «Ê·«"
 Else
 MsgBox "Please Eneter Area"
 End If
 Exit Sub
 End If
If val(.TextMatrix(row, .ColIndex("HaveDetail"))) = 2 And val(.TextMatrix(row, .ColIndex("EffectID"))) = 1 Then
If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
FrmDiviInvestmentCh.cmdOK.Enabled = True
'FrmDiviInvestmentCh.DcbType.Enabled = True
Else
FrmDiviInvestmentCh.cmdOK.Enabled = False
'FrmDiviInvestmentCh.DcbType.Enabled = False
End If
FrmDiviInvestmentCh.Ele(5).Caption = " ›«’Ì· " & " " & .TextMatrix(row, .ColIndex("Name")) & " " & "—ﬁ„ " & "  " & .TextMatrix(row, .ColIndex("PartNo"))
Load FrmDiviInvestmentCh
 FrmDiviInvestmentCh.show vbModal
 Else
 If SystemOptions.UserInterface = ArabicInterface Then
 MsgBox "·«Ì„ﬂ‰ ⁄—÷ ‘«‘…  ﬁ”Ì„«  «·»·Êﬂ «·« ›Ì Õ«·… «‰ ÌﬂÊ‰ «· ﬁ”Ì„ „ «Õ ··»Ì⁄  Ê·Â  ›«’Ì·"
 Else
 MsgBox "You can not display the divisions of the block except in the case to have Have details and  Available for sale"
 End If
End If
 End Select
End With
End Sub

Private Sub GridInstallments_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    With GridInstallments

        Select Case .ColKey(Col)
        Case "BlokNo"
        .ColComboList(.ColIndex("BlokNo")) = "..."
Case "Name"
  StrSQL = "select * from TblSpreading where UnitDetails=1 and Followed=" & val(DcbDivMain.BoundText) & ""
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
         
        End Select
    End With

End Sub

Private Sub ISButton3_Click()
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
ShowAttachments TxtSerial1.text, "170420167"
ErrTrap:
End Sub

Private Sub ISButton4_Click()
Dim i As Integer
With GridInstallments
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("DiviID"))) <> 0 Then
      If CheckSalesdET(val(.TextMatrix(i, .ColIndex("DiviID")))) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·⁄„·Ì… „— »ÿ… »›Ê« Ì— «·»Ì⁄"
    Else
    MsgBox "Can Not Delet This process linked to with invoices Sales"
    End If
    Exit Sub
    End If
    End If
    Next i
       .Clear flexClearScrollable, flexClearEverything
            .rows = 2
            
     End With
     RelinGrid
End Sub

Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double

    '---------------------- check if data Vaclete -----------------------
      If dcBranch.text = "" And val(dcBranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            dcBranch.SetFocus
            Exit Sub
     End If
           If DcbDivMain.text = "" And val(DcbDivMain.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·ÊÕœ…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please Select Unit ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            DcbDivMain.SetFocus
            Exit Sub
     End If
                If TxtCodeUnit.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  «œŒ«· «·ﬂÊœ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please enter code  ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            TxtCodeUnit.SetFocus
            Exit Sub
     End If
                If val(TxtDivArae.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡  «œŒ«· «·„”«Õ…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Please enter area ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
         End If
            TxtDivArae.SetFocus
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
     If val(DcbLand.BoundText) = 0 And DcbLand.text = "" Then
     If SystemOptions.UserInterface = ArabicInterface Then
     MsgBox "⁄›Ê«...«·—Ã«¡ ≈Œ Ì«— «·«—÷  "
     Else
     MsgBox "Please Select Land"
     End If
     DcbLand.SetFocus
     Exit Sub
     End If
Dim i As Integer
        With Me.GridInstallments
       For i = .FixedRows To .rows - 1
       If val(.TextMatrix(i, .ColIndex("TypeDivi"))) <> 0 Then
       If (.TextMatrix(i, .ColIndex("PartNo"))) = "" Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ ≈œŒ«· «·—ﬁ„ ›Ì «·”ÿ— —ﬁ„" & i
       Else
        MsgBox "Please Eneter No in Line" & i
       End If
       Exit Sub
       End If
              If val(.TextMatrix(i, .ColIndex("Area"))) = 0 Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ ≈œŒ«· «·„”«Õ… ›Ì «·”ÿ— —ﬁ„" & i
       Else
        MsgBox "Please Eneter Area in Line" & i
       End If
       Exit Sub
       End If
              If (.TextMatrix(i, .ColIndex("StraInform"))) = "" And val((.TextMatrix(i, .ColIndex("HaveDetail")))) = 2 And val((.TextMatrix(i, .ColIndex("EffectID")))) = 1 Then
       If SystemOptions.UserInterface = ArabicInterface Then
       MsgBox "Ì—ÃÏ ≈œŒ«· «· ›«’Ì· ›Ì «·”ÿ— —ﬁ„" & i
       Else
        MsgBox "Please Eneter details in Line" & i
       End If
       Exit Sub
       End If
      End If
     Next i
    End With

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
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
    Else
    MsgBox "Sorry Error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
    End If
End Sub
' new recored
'++++++++++++++++++++++++++++++++++++
Public Sub AddNewRec()
  'On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblDivInvesment", "ID", "")
    Me.TxtSerial1.text = StrRecID
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
    GridInstallments.rows = 1
sql = "SELECT     dbo.TblDivInvesmentDet.ID, dbo.TblDivInvesmentDet.DivInvID, dbo.TblDivInvesmentDet.Area, dbo.TblDivInvesmentDet.TotalArea, dbo.TblDivInvesmentDet.NewArea, "
sql = sql & "                       dbo.TblDivInvesmentDet.PartNo, dbo.TblDivInvesmentDet.BlokNo, dbo.TblDivInvesmentDet.Nourth, dbo.TblDivInvesmentDet.South, dbo.TblDivInvesmentDet.East,"
sql = sql & "                        dbo.TblDivInvesmentDet.West , dbo.TblDivInvesmentDet.EffectID, dbo.TblDivInvesmentDet.TypeDivi, dbo.TblSpreading.name, dbo.TblSpreading.NameE ,dbo.TblDivInvesmentDet.StraInform ,dbo.TblDivInvesmentDet.HaveDetail ,dbo.TblDivInvesmentDet.TypBP "
sql = sql & "   FROM         dbo.TblDivInvesmentDet LEFT OUTER JOIN"
sql = sql & "                        dbo.TblSpreading ON dbo.TblDivInvesmentDet.TypeDivi = dbo.TblSpreading.ID"
sql = sql & "   Where (dbo.TblDivInvesmentDet.DivInvID =" & val(TxtSerial1.text) & ") "

  Rs1.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim i As Integer
     With Me.GridInstallments
                    For i = .FixedRows To Rs1.RecordCount
                   .rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(i, .ColIndex("Ser")) = i
                   .TextMatrix(i, .ColIndex("DiviID")) = IIf(IsNull(Rs1("ID").value), 0, Rs1("ID").value)
                   .TextMatrix(i, .ColIndex("TypBP")) = IIf(IsNull(Rs1("TypBP").value), 0, Rs1("TypBP").value)
                   .TextMatrix(i, .ColIndex("HaveDetail")) = IIf(IsNull(Rs1("HaveDetail").value), 1, Rs1("HaveDetail").value)
                   .TextMatrix(i, .ColIndex("TypeDivi")) = IIf(IsNull(Rs1("TypeDivi").value), 0, Rs1("TypeDivi").value)
                   .TextMatrix(i, .ColIndex("Area")) = IIf(IsNull(Rs1("Area").value), "", Rs1("Area").value)
                   .TextMatrix(i, .ColIndex("TotalArea")) = IIf(IsNull(Rs1("TotalArea").value), "", Rs1("TotalArea").value)
                   .TextMatrix(i, .ColIndex("NewArea")) = IIf(IsNull(Rs1("NewArea").value), "", Rs1("NewArea").value)
                   .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(Rs1("PartNo").value), "", Rs1("PartNo").value)
                   .TextMatrix(i, .ColIndex("BlokNo")) = IIf(IsNull(Rs1("BlokNo").value), "", Rs1("BlokNo").value)
                   .TextMatrix(i, .ColIndex("EffectID")) = IIf(IsNull(Rs1("EffectID").value), 2, Rs1("EffectID").value)
                   .TextMatrix(i, .ColIndex("Nourth")) = IIf(IsNull(Rs1("Nourth").value), "", Rs1("Nourth").value)
                   .TextMatrix(i, .ColIndex("South")) = IIf(IsNull(Rs1("South").value), "", Rs1("South").value)
                   .TextMatrix(i, .ColIndex("East")) = IIf(IsNull(Rs1("East").value), "", Rs1("East").value)
                   .TextMatrix(i, .ColIndex("West")) = IIf(IsNull(Rs1("West").value), "", Rs1("West").value)
                   .TextMatrix(i, .ColIndex("StraInform")) = IIf(IsNull(Rs1("StraInform").value), "", Rs1("StraInform").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next i
        End With
        Exit Sub
ErrTrap:
    End Sub



Private Sub DcbLand_Change()
Dim Fullcode As String
Dim Area As Double
If val(DcbLand.BoundText) <> 0 Then
GetTblBuyLandRealEstate val(DcbLand.BoundText), Fullcode, 0
Me.Text1.text = Fullcode
End If
If Me.TxtModFlg.text <> "R" Then
GetLandInformation val(DcbLand.BoundText), Area
TxtAlwArea.text = Area
TxtAlwArea.text = val(TxtAlwArea.text) - SumValuTotal(val(TxtSerial1.text), val(DcbLand.BoundText), 1)
End If
End Sub

Private Sub DcbLand_Click(Area As Integer)
DcbLand_Change
End Sub

Private Sub ISButton6_Click()
   With Me.GridInstallments
   If .rows < 2 Then Exit Sub
           If CheckSalesdET(val(.TextMatrix(.row, .ColIndex("DiviID")))) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·⁄„·Ì… „— »ÿ… »›Ê« Ì— «·»Ì⁄"
    Else
    MsgBox "Can Not Delet This process linked to with invoices Sales"
    End If
    Exit Sub
    End If
   .RemoveItem .row
   End With
   RelinGrid
End Sub

Private Sub ISButton8_Click()
FrmSearchinvestment.inde = 12
Load FrmSearchinvestment
FrmSearchinvestment.show vbModal
End Sub

Private Sub lbl_Change(index As Integer)
lbl_Click index
End Sub

Private Sub lbl_Click(index As Integer)
Select Case index
Case 6
TxtAlwAreaAfter.text = val(TxtAlwArea.text) - val(lbl(6).Caption)
End Select
End Sub



Private Sub RdType_Click(index As Integer)
If Me.TxtModFlg.text <> "R" Then
Dim Dcombos As New ClsDataCombos
If RdType(0).value = True Then
    Dcombos.GetLandActive DcbLand
DcbInvise.Enabled = True
 Else
 DcbInvise.BoundText = 0
DcbInvise.Enabled = False
Dcombos.GetLandNotActive DcbLand
 End If
 Else
 Dcombos.GetLand DcbLand
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim ID  As Integer
GetTblBuyLandRealEstate ID, Me.Text1.text, 1
DcbLand.BoundText = ID
End Sub




Private Sub TxtAlwArea_Change()
If Me.TxtModFlg.text <> "R" Then
TxtAlwAreaAfter.text = val(TxtAlwArea.text) - val(lbl(6).Caption)
End If
End Sub

Private Sub TxtAlwAreaAfter_Change()
If Me.TxtModFlg.text <> "R" Then
If val(TxtAlwAreaAfter.text) <> 0 Then
TxtMetreValue.text = val(TxtCurrValue.text) / val(TxtAlwAreaAfter.text)
TxtMetreValue.text = Round(val(TxtMetreValue.text), 2)
End If
If val(TxtSharNo.text) <> 0 Then
TxtSharMetre.text = val(TxtAlwAreaAfter.text) / val(TxtSharNo.text)
TxtSharMetre = Round(val(TxtSharMetre.text), 2)
End If
End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
Me.DcbInvise.BoundText = TxtCode.text
End Sub

Private Sub TxtCurrValue_Change()
lbl(15).Caption = WriteNo(val(Me.TxtCurrValue.text), 0)
If Me.TxtModFlg.text <> "R" Then
If val(TxtAlwAreaAfter.text) <> 0 Then
TxtMetreValue.text = val(TxtCurrValue.text) / val(TxtAlwAreaAfter.text)
TxtMetreValue.text = Round(val(TxtMetreValue.text), 2)
End If
End If
End Sub

Private Sub TxtDivArae_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtDivArae.text, 0)
End Sub

Private Sub TxtDivArae_LostFocus()
If Me.TxtModFlg.text <> "R" Then
If Round(val(TxtDivArae.text), 2) > Round(val(TxtAlwArea.text), 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·„”«Õ… «ﬂ»— „‰ «·„”«Õ… «·„ «Õ… ·· ﬁ”Ì„"
Else
MsgBox "The area can not larger than total arae"
End If
TxtDivArae.text = 0
Exit Sub
End If
End If
End Sub

' change id search
Private Sub TxtSerial1_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub
Sub GetActiveInvestInformation(Optional ID As Double = 0)
If ID <> 0 Then
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
sql = "Select * from TblActivateInvestment  where InviseNo=" & ID & ""
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
Rs4.MoveFirst
TxtAlwArea.text = IIf(IsNull(Rs4("Area").value), 0, Rs4("Area").value)
Else
TxtAlwArea.text = 0
End If
End If
End Sub


Function SumValuTotal(Optional ID As Double = 0, Optional InvesID As Double = 0, Optional Typ As Integer = 0) As Double
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
Dim sql As String
SumValuTotal = 0
sql = "SELECT     SUM(Total) AS sumValu"
sql = sql & " From dbo.TblDivInvesment"
If Typ = 0 Then
sql = sql & " Where (id <> " & ID & ") And (InvesID = " & InvesID & ")"
Else
sql = sql & " Where (id <> " & ID & ") And (LandID = " & InvesID & ")"
End If
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
SumValuTotal = IIf(IsNull(Rs4("sumValu").value), 0, Rs4("sumValu").value)
Else
SumValuTotal = 0
End If
End Function

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
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim sql As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim i As Integer
    Dim ID As Double
        If CheckBuy(val(DcbInvise.BoundText)) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·⁄„·Ì… „— »ÿ…  »«· ‰«“·"
    Else
    MsgBox "Can Not Delete This process linked to with invoices Buy"
    End If
    Exit Sub
    End If
    
    
    If CheckSales(val(DcbInvise.BoundText)) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰ «·Õ–› Â–Â «·⁄„·Ì… „— »ÿ… »›Ê« Ì— «·»Ì⁄"
    Else
    MsgBox "Can Not Delet This process linked to with invoices Sales"
    End If
    Exit Sub
    End If
    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
       End If
               Else
                sql = "Update TblExpensesInvesment set DivPayed=Null where InvesID=" & val(DcbInvise.BoundText) & " "
                Cn.Execute sql
                StrSQL = "delete From TblDivInvestInformation  where  DivID =" & val(TxtSerial1.text) & ""
                   Cn.Execute StrSQL, , adExecuteNoRecords
                      StrSQL = "delete From TblDivInvesmentDetCheld  where  DivID =" & val(TxtSerial1.text) & ""
                   Cn.Execute StrSQL, , adExecuteNoRecords
                   
                RsSavRec.Find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
                  StrSQL = "Delete From TblDivInvesmentDet Where DivInvID =" & val(TxtSerial1.text) & ""
                 Cn.Execute StrSQL, , adExecuteNoRecords
                                          RsSavRec.delete
                                
                                    
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
             
               '''''''''''''''''''''''''''''''

                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title)
               End If

     End If                       '------------------------------ Move Next ---------------------------.
        Me.Refresh
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
    ElseIf TxtModFlg.text = "R" Then
     XPDtbTrans.Enabled = False
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
    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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
    If TxtSerial1.text <> "" Then
'     If CheckSales(val(DcbInvise.BoundText)) = True Then
'    If SystemOptions.UserInterface = ArabicInterface Then
'    MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·⁄„·Ì… „— »ÿ… »›Ê« Ì— «·»Ì⁄"
'    Else
'    MsgBox "Can Not Update This process linked to with invoices Sales"
'    End If
'    Exit Sub
'    End If
     If CheckBuy(val(DcbInvise.BoundText)) = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "·«Ì„ﬂ‰  ⁄œÌ· Â–Â «·⁄„·Ì… „— »ÿ…  »«· ‰«“·"
    Else
    MsgBox "Can Not Update This process linked to with invoices Buy"
    End If
    Exit Sub
    End If
    
        TxtModFlg = "E"
            GridInstallments.rows = GridInstallments.rows + 1
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
    RdType(0).value = True
    lbl(6).Caption = 0
    TxtModFlg.text = "N"
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.rows = 2
    Me.DCboUserName.BoundText = user_id
    Me.dcBranch.BoundText = Current_branch
    dcBranch.SetFocus
  
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
  MySQL = "SELECT     dbo.TblDivInvesment.ID, dbo.TblDivInvesment.RecordDate, dbo.TblDivInvesment.CurrValue, dbo.TblDivInvesment.DevlopValue, "
  MySQL = MySQL & "                   dbo.TblDivInvesment.AfterDevlopValue, dbo.TblDivInvesment.ShareValue, dbo.TblDivInvesment.SharNo, dbo.TblDivInvesment.Remarks, dbo.TblDivInvesment.Total,"
  MySQL = MySQL & "                     dbo.TblDivInvesment.AlwArea, dbo.TblDivInvesment.AlwAreaAfter, dbo.TblDivInvesment.BranchID, dbo.TblBranchesData.branch_name,"
  MySQL = MySQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblDivInvesment.InvesID, dbo.Tblinvestment.Name, dbo.Tblinvestment.NameE, dbo.TblDivInvesment.LandID,"
  MySQL = MySQL & "                     dbo.TblBuyLanReEst.Name AS LandName, dbo.TblBuyLanReEst.NameE AS LandNameE, dbo.TblDivInvesmentDet.TypeDivi, dbo.TblSpreading.Name AS DivName,"
  MySQL = MySQL & "                     dbo.TblSpreading.NameE AS DivNameE, dbo.TblDivInvesmentDet.EffectID, dbo.TblDivInvesmentDet.Area, dbo.TblDivInvesmentDet.TotalArea,"
  MySQL = MySQL & "                     dbo.TblDivInvesmentDet.NewArea, dbo.TblDivInvesmentDet.PartNo, dbo.TblDivInvesmentDet.BlokNo, dbo.TblDivInvesmentDet.Nourth,"
  MySQL = MySQL & "                     dbo.TblDivInvesmentDet.South , dbo.TblDivInvesmentDet.East, dbo.TblDivInvesmentDet.West"
  MySQL = MySQL & "   FROM         dbo.TblSpreading RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblDivInvesmentDet ON dbo.TblSpreading.ID = dbo.TblDivInvesmentDet.TypeDivi RIGHT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblDivInvesment ON dbo.TblDivInvesmentDet.DivInvID = dbo.TblDivInvesment.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblBuyLanReEst ON dbo.TblDivInvesment.LandID = dbo.TblBuyLanReEst.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.Tblinvestment ON dbo.TblDivInvesment.InvesID = dbo.Tblinvestment.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblBranchesData ON dbo.TblDivInvesment.BranchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "  Where (dbo.TblDivInvesment.ID =" & val(TxtSerial1.text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDivInvestment.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDivInvestmentE.rpt"
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


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Divide  Land  "
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
    lbl(9).Caption = "Area Available "
    lbl(11).Caption = "Meter Value"
    lbl(10).Caption = "Area Available After"
    Frame6.Caption = "Data"
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
  .TextMatrix(0, .ColIndex("Name")) = "Type Division"
  .TextMatrix(0, .ColIndex("Nourth")) = "North"
  .TextMatrix(0, .ColIndex("South")) = "South"
  .TextMatrix(0, .ColIndex("East")) = "East"
  .TextMatrix(0, .ColIndex("West")) = "West"
  .TextMatrix(0, .ColIndex("Area")) = "Area"
  .TextMatrix(0, .ColIndex("EffectID")) = "Effect"
  .TextMatrix(0, .ColIndex("TotalArea")) = "Total Area"
   .TextMatrix(0, .ColIndex("NewArea")) = "New Area"
  .TextMatrix(0, .ColIndex("BlokNo")) = "Block No."
  .TextMatrix(0, .ColIndex("PartNo")) = "Part No."
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
For i = 1 To .rows - 1
If val(.TextMatrix(i, .ColIndex("DiviID"))) <> 0 Then
Counter = Counter + 1
.TextMatrix(i, .ColIndex("Ser")) = Counter
'If val(.TextMatrix(I, .ColIndex("EffectID"))) = 2 Then
summation = summation + val(.TextMatrix(i, .ColIndex("Area")))
'End If
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
   My_SQL = "TblDivInvesment"
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
Private Sub TxtSharNo_Change()
If Me.TxtModFlg.text <> "R" Then
If val(TxtSharNo.text) <> 0 Then
TxtSharMetre.text = val(TxtAlwAreaAfter.text) / val(TxtSharNo.text)
TxtSharMetre = Round(val(TxtSharMetre.text), 2)
End If
End If
End Sub
