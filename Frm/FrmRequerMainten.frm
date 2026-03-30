VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmRequerMainten 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»«  «·’Ì«‰…"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17070
   Icon            =   "FrmRequerMainten.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   17070
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   7695
      Left            =   -360
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   600
      Width           =   17415
      Begin VB.ComboBox DcbTypeMaint 
         Height          =   315
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·ÿ·»"
         Height          =   7035
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   600
         Width           =   16905
         Begin VB.TextBox supervisorNotes 
            Alignment       =   1  'Right Justify
            Height          =   675
            Left            =   8520
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   133
            Top             =   3720
            Width           =   7335
         End
         Begin VB.TextBox NoOfLabs 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8520
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   3240
            Width           =   2655
         End
         Begin VB.TextBox tripKM 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3120
            TabIndex        =   127
            Top             =   6240
            Width           =   1335
         End
         Begin VB.CheckBox UpdateKM 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕœÌÀ ﬁ—«¡Â «·⁄œ«œ"
            Height          =   255
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   6240
            Value           =   1  'Checked
            Width           =   2295
         End
         Begin VB.TextBox DifferentKm 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   125
            Top             =   5880
            Width           =   1335
         End
         Begin VB.TextBox ManualKM 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3120
            TabIndex        =   123
            Top             =   5880
            Width           =   1335
         End
         Begin VB.TextBox LastKM 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5850
            Locked          =   -1  'True
            TabIndex        =   121
            Top             =   5880
            Width           =   1335
         End
         Begin VB.TextBox RemainKmToArrive 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5280
            TabIndex        =   118
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox StopLocation 
            Alignment       =   1  'Right Justify
            Height          =   915
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   116
            Top             =   840
            Width           =   2775
         End
         Begin VB.ComboBox EquipmentStatusid 
            Height          =   315
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton showAll 
            Caption         =   "⁄—÷ «·ﬂ·"
            Height          =   375
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   6480
            Width           =   1575
         End
         Begin VB.TextBox TxtBoardNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8520
            TabIndex        =   95
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox TxtRejecReason 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8520
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   94
            Top             =   4440
            Width           =   7335
         End
         Begin VB.ComboBox DcbStatusMaint 
            Height          =   315
            ItemData        =   "FrmRequerMainten.frx":038A
            Left            =   12720
            List            =   "FrmRequerMainten.frx":038C
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   3240
            Width           =   3135
         End
         Begin VB.TextBox TxtOperationNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6240
            TabIndex        =   90
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox TxtExternaExam 
            Alignment       =   1  'Right Justify
            Height          =   795
            Left            =   3960
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   86
            Top             =   3600
            Width           =   3135
         End
         Begin VB.TextBox TxtRemarksEqup 
            Alignment       =   1  'Right Justify
            Height          =   795
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   3600
            Width           =   2895
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·„ «·„⁄œ…"
            Height          =   1095
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   1920
            Width           =   8535
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox TxtDrievName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   600
               Width           =   6855
            End
            Begin XtremeSuiteControls.RadioButton dcbDrievType 
               Height          =   255
               Index           =   0
               Left            =   6840
               TabIndex        =   82
               Top             =   240
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊŸ›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDrievID 
               Bindings        =   "FrmRequerMainten.frx":038E
               Height          =   315
               Left            =   240
               TabIndex        =   83
               Top             =   240
               Width           =   5775
               _ExtentX        =   10186
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
            Begin XtremeSuiteControls.RadioButton dcbDrievType 
               Height          =   255
               Index           =   1
               Left            =   6840
               TabIndex        =   84
               Top             =   600
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "€Ì— „ÊŸ›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   14430
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   600
            Width           =   945
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁ«∆œ «·„⁄œ…"
            Height          =   975
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   960
            Width           =   8535
            Begin VB.TextBox TxtMobile 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               TabIndex        =   112
               Top             =   600
               Width           =   1815
            End
            Begin VB.TextBox TxtLeaderName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   600
               Width           =   4215
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   240
               Width           =   1065
            End
            Begin XtremeSuiteControls.RadioButton dcbLeaderType 
               Height          =   255
               Index           =   0
               Left            =   6840
               TabIndex        =   72
               Top             =   240
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊŸ›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbLeaderID 
               Bindings        =   "FrmRequerMainten.frx":03A3
               Height          =   315
               Left            =   120
               TabIndex        =   73
               Top             =   240
               Width           =   5895
               _ExtentX        =   10398
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
            Begin XtremeSuiteControls.RadioButton dcbLeaderType 
               Height          =   255
               Index           =   1
               Left            =   6840
               TabIndex        =   78
               Top             =   600
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "€Ì— „ÊŸ›"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ÃÊ«·"
               Height          =   285
               Index           =   24
               Left            =   1680
               TabIndex        =   113
               Top             =   600
               Width           =   1125
            End
         End
         Begin VB.TextBox TxtEnterCounter 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3240
            TabIndex        =   67
            Top             =   1920
            Width           =   3855
         End
         Begin VB.TextBox TxtOutCounter 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox TxtProblemOther 
            Alignment       =   1  'Right Justify
            Height          =   435
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   240
            Width           =   2775
         End
         Begin VB.ComboBox ProblemTimID 
            Height          =   315
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   240
            Width           =   1575
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ ÊÊﬁ  «·„‘ﬂ·…"
            Height          =   1155
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   2280
            Width           =   6975
            Begin MSComCtl2.DTPicker StartDate 
               Height          =   315
               Left            =   3480
               TabIndex        =   4
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144703489
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker EndExptedDate 
               Height          =   315
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144703489
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker EndActTIme 
               Height          =   315
               Left            =   120
               TabIndex        =   7
               Top             =   720
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144703490
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker StartTime 
               Height          =   315
               Left            =   3480
               TabIndex        =   5
               Top             =   720
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   556
               _Version        =   393216
               Format          =   144703490
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·⁄„·"
               Height          =   285
               Index           =   9
               Left            =   5520
               TabIndex        =   58
               Top             =   255
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «· Êﬁ›"
               Height          =   285
               Index           =   12
               Left            =   1710
               TabIndex        =   57
               Top             =   255
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”«⁄Â «·⁄„·"
               Height          =   285
               Index           =   15
               Left            =   4950
               TabIndex        =   56
               Top             =   720
               Width           =   1605
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Êﬁ  «·ÿ·»"
               Height          =   285
               Index           =   16
               Left            =   1680
               TabIndex        =   55
               Top             =   720
               Width           =   1605
            End
         End
         Begin VB.TextBox TxtDes 
            Alignment       =   1  'Right Justify
            Height          =   555
            Left            =   8520
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   4800
            Width           =   7455
         End
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   1275
            Left            =   8520
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   5520
            Width           =   7335
         End
         Begin MSDataListLib.DataCombo DcbEquepment 
            Height          =   315
            Left            =   11520
            TabIndex        =   3
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbBranchIDTo 
            Bindings        =   "FrmRequerMainten.frx":03B8
            Height          =   315
            Left            =   3000
            TabIndex        =   64
            Top             =   1080
            Width           =   3735
            _ExtentX        =   6588
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
         Begin MSDataListLib.DataCombo DcbOperiatorID 
            Bindings        =   "FrmRequerMainten.frx":03CD
            Height          =   315
            Left            =   11520
            TabIndex        =   75
            Top             =   600
            Width           =   2895
            _ExtentX        =   5106
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   315
            Left            =   6240
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   600
            Width           =   2205
            _cx             =   3889
            _cy             =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
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
            Begin VB.TextBox txtLetter1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1935
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   0
               Width           =   285
            End
            Begin VB.TextBox txtLetter2 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1710
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   0
               Width           =   240
            End
            Begin VB.TextBox txtLetter3 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1440
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   0
               Width           =   315
            End
            Begin VB.TextBox txtNum1 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   795
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   0
               Width           =   360
            End
            Begin VB.TextBox txtNum2 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   480
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   0
               Width           =   330
            End
            Begin VB.TextBox txtNum3 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   270
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   0
               Width           =   300
            End
            Begin VB.TextBox txtLetter4 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1155
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   0
               Width           =   360
            End
            Begin VB.TextBox txtNum4 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   0
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   0
               Width           =   300
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid13 
            Height          =   1035
            Left            =   120
            TabIndex        =   106
            Top             =   4560
            Width           =   8280
            _cx             =   14605
            _cy             =   1826
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmRequerMainten.frx":03E2
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
         Begin ImpulseButton.ISButton removeRow 
            Height          =   420
            Left            =   7320
            TabIndex        =   109
            Top             =   5520
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› ”ÿ— "
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
            ButtonImage     =   "FrmRequerMainten.frx":0490
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton clearGridBtn 
            Height          =   420
            Left            =   6000
            TabIndex        =   110
            Top             =   5520
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–› «·ﬂ·"
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
            ButtonImage     =   "FrmRequerMainten.frx":0A2A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker expectedEndDate 
            Height          =   315
            Left            =   5400
            TabIndex        =   129
            Top             =   6600
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   138412033
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker expectedEndtime 
            Height          =   315
            Left            =   3960
            TabIndex        =   131
            Top             =   6600
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   138412034
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·„—«ﬁ»"
            Height          =   405
            Index           =   34
            Left            =   16080
            TabIndex        =   136
            Top             =   3720
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·—œÊœ "
            Height          =   285
            Index           =   32
            Left            =   11400
            TabIndex        =   134
            Top             =   3240
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Œ—ÊÃ «·„ Êﬁ⁄"
            Height          =   285
            Index           =   31
            Left            =   6960
            TabIndex        =   130
            Top             =   6615
            Width           =   1485
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œ«œ «· ‘€Ì·"
            Height          =   255
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   6240
            Width           =   1455
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—ﬁ"
            Height          =   255
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   5880
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬂÌ·Ê „ — «·Õ«·Ì"
            Height          =   255
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   5880
            Width           =   1455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬂÌ·Ê „ — «·”«»ﬁ"
            Height          =   255
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   5880
            Width           =   1935
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„ﬂ«‰ «·⁄ÿ·"
            Height          =   255
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„”«›… «·„ »ﬁÌ… ··Ê’Ê·"
            Height          =   255
            Left            =   6360
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·Â «·„⁄œ…"
            Height          =   285
            Index           =   26
            Left            =   4560
            TabIndex        =   115
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„·Õﬁ« "
            Height          =   285
            Index           =   23
            Left            =   6960
            TabIndex        =   107
            Top             =   4320
            Width           =   1245
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «··ÊÕ…"
            Height          =   285
            Index           =   22
            Left            =   10320
            TabIndex        =   96
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «·—›÷"
            Height          =   405
            Index           =   21
            Left            =   15720
            TabIndex        =   93
            Top             =   4440
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·ÿ·»"
            Height          =   285
            Index           =   20
            Left            =   15600
            TabIndex        =   92
            Top             =   3240
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «· ‘€Ì·"
            Height          =   285
            Index           =   19
            Left            =   10320
            TabIndex        =   89
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›Õ’ «·Œ«—ÃÌ"
            Height          =   405
            Index           =   17
            Left            =   7200
            TabIndex        =   88
            Top             =   3720
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  ⁄·Ï «·„⁄œ…"
            Height          =   405
            Index           =   14
            Left            =   3120
            TabIndex        =   87
            Top             =   3720
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„—«ﬁ» «· ‘€Ì·    "
            Height          =   255
            Left            =   15240
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œ«œ «·œŒÊ·"
            Height          =   285
            Index           =   13
            Left            =   6960
            TabIndex        =   69
            Top             =   1920
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œ«œ «·Œ—ÊÃ"
            Height          =   285
            Index           =   10
            Left            =   2040
            TabIndex        =   68
            Top             =   1920
            Width           =   1125
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄ ÿ«·» «·«’·«Õ"
            Height          =   255
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Êﬁ  «·„‘ﬂ·…"
            Height          =   285
            Index           =   3
            Left            =   4680
            TabIndex        =   59
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄ÿ·"
            Height          =   285
            Index           =   2
            Left            =   15960
            TabIndex        =   53
            Top             =   4920
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„⁄œÂ"
            Height          =   285
            Index           =   29
            Left            =   15240
            TabIndex        =   52
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·›‰Ì"
            Height          =   645
            Index           =   18
            Left            =   16080
            TabIndex        =   49
            Top             =   5880
            Width           =   645
         End
      End
      Begin VB.ComboBox Contract_period 
         Height          =   315
         ItemData        =   "FrmRequerMainten.frx":0FC4
         Left            =   18840
         List            =   "FrmRequerMainten.frx":0FCE
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox XPTxtID 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   14850
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmRequerMainten.frx":0FDC
         Height          =   315
         Left            =   6000
         TabIndex        =   42
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
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
      Begin MSComCtl2.DTPicker XPDtbTrans 
         Height          =   315
         Left            =   12240
         TabIndex        =   50
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   138412033
         CurrentDate     =   41640
      End
      Begin MSDataListLib.DataCombo DcbUnit 
         Bindings        =   "FrmRequerMainten.frx":0FF1
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         Height          =   255
         Index           =   0
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Â–… «·‘«‘…  ﬁÊ„ » ”ÃÌ· ÿ·»«  «·’Ì«‰…"
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
         Height          =   375
         Index           =   25
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   -480
         Width           =   3735
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   495
         Left            =   480
         Top             =   -600
         Width           =   3855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   -1320
         TabIndex        =   47
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· «—ÌŒ"
         Height          =   285
         Index           =   1
         Left            =   13590
         TabIndex        =   45
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÿ·»"
         Height          =   285
         Index           =   4
         Left            =   16110
         TabIndex        =   44
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblbr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄ «·ﬁ«∆„ »«·«’·«Õ"
         Height          =   255
         Left            =   10440
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   13410
      TabIndex        =   36
      Top             =   750
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox TxtNoteSerial 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   345
      Left            =   14310
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox TxtNoteSerial1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Left            =   14190
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox oldtxtNoteSerial1 
      Height          =   285
      Left            =   19470
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtNoteID 
      Height          =   285
      Left            =   14190
      TabIndex        =   32
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   17055
      _cx             =   30083
      _cy             =   1032
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "ÿ·»«  «·’Ì«‰…  "
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   375
         Index           =   0
         Left            =   1185
         TabIndex        =   11
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
         ButtonImage     =   "FrmRequerMainten.frx":1006
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
         Left            =   120
         TabIndex        =   12
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
         ButtonImage     =   "FrmRequerMainten.frx":13A0
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
         Left            =   1710
         TabIndex        =   13
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
         ButtonImage     =   "FrmRequerMainten.frx":173A
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
         Left            =   645
         TabIndex        =   14
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
         ButtonImage     =   "FrmRequerMainten.frx":1AD4
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   3960
         Picture         =   "FrmRequerMainten.frx":1E6E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2400
         TabIndex        =   30
         Top             =   0
         Width           =   2205
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic4 
      Height          =   540
      Left            =   4110
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8760
      Width           =   9225
      _cx             =   16272
      _cy             =   953
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   0
         Left            =   8310
         TabIndex        =   16
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   1
         Left            =   7455
         TabIndex        =   17
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   2
         Left            =   6615
         TabIndex        =   18
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   3
         Left            =   5760
         TabIndex        =   19
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   4905
         TabIndex        =   20
         Top             =   75
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   6
         Left            =   1080
         TabIndex        =   21
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   375
         Left            =   1935
         TabIndex        =   22
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   5
         Left            =   3840
         TabIndex        =   29
         Top             =   60
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
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   9
         Left            =   3000
         TabIndex        =   31
         Top             =   60
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   8
         Left            =   0
         TabIndex        =   111
         Top             =   60
         Width           =   1005
         _ExtentX        =   1773
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
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   315
      Left            =   12060
      TabIndex        =   23
      Top             =   8400
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   13470
      TabIndex        =   37
      Top             =   3570
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   7
      Left            =   13830
      TabIndex        =   38
      Top             =   1920
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·ÿ·»"
      Height          =   285
      Index           =   33
      Left            =   0
      TabIndex        =   135
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·‘Œ’ «·„”∆Ê·"
      Height          =   285
      Index           =   28
      Left            =   4080
      TabIndex        =   51
      Top             =   2160
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÌœ:"
      Height          =   315
      Index           =   30
      Left            =   13080
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Õ—— »Ê«”ÿ… : "
      Height          =   270
      Index           =   8
      Left            =   15885
      TabIndex        =   28
      Top             =   8355
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " «·”Ã· «·Õ«·Ì:"
      Height          =   315
      Index           =   7
      Left            =   4950
      TabIndex        =   27
      Top             =   8460
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ⁄œœ «·”Ã·« :"
      Height          =   315
      Index           =   6
      Left            =   3210
      TabIndex        =   26
      Top             =   8460
      Width           =   975
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   2490
      TabIndex        =   25
      Top             =   8460
      Width           =   495
   End
   Begin VB.Label XPTxtCurrent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   4260
      TabIndex        =   24
      Top             =   8460
      Width           =   615
   End
End
Attribute VB_Name = "FrmRequerMainten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String
Dim rs_CarParts As ADODB.Recordset
Public bol As Boolean
Public novalue As Boolean
Dim ODERdATEFocus As Boolean
Dim ODERTimeFocus As Boolean


'Private Sub Accredit_Click()
'    Dim BeginTrans As Boolean
'
'    Cn.BeginTrans
'    BeginTrans = True
'
'    If IsNull(rs("Posted")) Then
'        rs("Posted") = user_id
'        rs("PostedDate") = Time
'    Else
'        rs("Posted") = Null
'       rs("PostedDate") = Time
'    End If
'
'    rs.update
' If SystemOptions.UserInterface = ArabicInterface Then
'    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
'Else
'Accredit.Caption = "Sent To approval "
'End If

'    Cn.CommitTrans
'    BeginTrans = False
'FillApprovedTable
'    Retrive (val(Me.XPTxtID.text))
'End Sub

'Private Sub bClose_Click()
'Frame6.Visible = False
'If Me.ChekAccept.value = xtpChecked Then
'Frame2.Visible = True
'End If
'If Me.ChekContracted.value = xtpChecked Then
'Frame5.Visible = True
'End If
'End Sub

'Private Sub ChekAccept_Click()
'If Me.ChekAccept.value = vbChecked Then
'Me.CHekNotAccept.value = vbUnchecked
'Me.ChekContracted.value = vbUnchecked
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'Me.Frame2.Visible = True
'Me.Frame5.Visible = False
'Else
'Me.Frame2.Visible = False
'End If
'End Sub
'Private Sub RemoveGridRow()
'
'    With Me.Fg
'
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'
'    ReLineGrid
'End Sub
'Private Sub RemoveGridRow2()
'
'    With Me.fg2
'
'        If .Row <= 0 Then Exit Sub
'        .RemoveItem .Row
'    End With
'
'    ReLineGrid
'End Sub

'Private Sub ChekContracted_Click()
'If Me.ChekContracted.value = xtpChecked Then
'Me.CHekNotAccept.value = xtpUnchecked
'Me.ChekAccept.value = xtpUnchecked
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'Me.Frame2.Visible = False
'Frame5.Visible = True
'Else
'Me.Frame5.Visible = False
'End If
'
'End Sub

'Private Sub CHekNotAccept_Click()
'If Me.CHekNotAccept.value = vbChecked Then
'Me.Frame2.Visible = False
'Me.Frame5.Visible = False
'lbl(36).Visible = True
'Me.txtnotAccept.Visible = True
'Me.ChekAccept.value = vbUnchecked
''Me.ChekContracted.value = vbUnchecked
'Else
'Me.Frame2.Visible = True
'lbl(36).Visible = False
'Me.txtnotAccept.Visible = False
'End If
'End Sub
Function VIEW_ATTACH()
    'On Error Resume Next
 
    'If TxtEmp_Code.text = "" Then MsgBox "·«»œ „‰ «Õ Ì«— „ÊŸ› «Ê·«": Exit Sub

    imaged.show
    imaged.Label9.Caption = "„—›ﬁ«  ÿ·» ’Ì«‰Â —ﬁ„"
    imaged.Caption = "„—›ﬁ«  ÿ·» ’Ì«‰Â  "
    imaged.txtopeation_type = "ÿ·» ’Ì«‰Â"
    imaged.SUBJECT_NO = XPTxtID.Text  'TxtEmp_Code.text
    imaged.Label6.Caption = "—ﬁ„ «·ÿ·»"
    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type = 'ÿ·» ’Ì«‰Â' and subject_no='" & XPTxtID.Text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Function



Private Sub Check1_Click()

End Sub

Private Sub Cmd_Click(Index As Integer)

    ' On Error GoTo ErrTrap
    Select Case Index

        
        Case 0
          

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"

            clear_all Me
      dcbLeaderType_Click (0)
            dcbLeaderType(0).value = True
            dcbDrievType_Click (0)
            dcbDrievType(0).value = True
            DcbStatusMaint.ListIndex = 0
           DcbTypeMaint.ListIndex = 0
           DcbStatusMaint.ListIndex = 0
 'EquipmentStatusid.ListIndex = 0
ODERdATEFocus = False
ODERTimeFocus = False


            Me.DCboUserName.BoundText = user_id
        '    TxtPaymentCounts.text = 1
dcBranch.BoundText = Current_branch


        Case 1
If val(DcbStatusMaint.ListIndex) = 2 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«Ì„ﬂ‰ «· ⁄œÌ· .·ﬁœ  „ «·«‰ Â«¡ „‰ ÿ·» «·’Ì«‰Â"
Else
MsgBox "You can note edit.This is process completed"
End If
Exit Sub
End If

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
'            Me.DCboUserName.BoundText = user_id

        Case 2
    
            Dim Msg As String


If Me.TxtModFlg = "N" Then
If ODERdATEFocus = False Then
 
                    If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "·«»œ „‰ ﬂ «»…  «—ÌŒ «·ÿ·»"
                            Else
                                Msg = "íMust enter  Start Work time"
                            End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
  End If
  
  
  If ODERTimeFocus = False Then
 
                    If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "·«»œ „‰ ﬂ «»… Êﬁ  «·ÿ·»"
                            Else
                                Msg = "íMust enter  Start Work time"
                            End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
  End If
  
  If val(ManualKM.Text) = 0 Then
 
                    If SystemOptions.UserInterface = ArabicInterface Then
                                Msg = "·«»œ „‰ ﬂ «»… «·⁄œ«œ   «·Õ«·Ì  "
                            Else
                                Msg = "íMust enter Counter"
                            End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Screen.MousePointer = vbDefault
                Exit Sub
       
       
       
End If

End If



            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Branch Based Reform"
                Else
                    Msg = "Õœœ «·›—⁄ «·ﬁ«∆„ »«·«’·«Õ "
                End If
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
  If val(DcbBranchIDTo.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«—   «·›—⁄ «·ÿ«·» ··’Ì«‰…"
Else
MsgBox "Please Select Branch Request Maintenance"
End If
DcbBranchIDTo.SetFocus
Exit Sub
End If



If val(EquipmentStatusid.ListIndex) = -1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«—  Õ«·… «·„⁄œ…"
Else
MsgBox "Please Select Order Status"
End If
EquipmentStatusid.SetFocus
Exit Sub
End If


If val(DcbStatusMaint.ListIndex) = -1 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«—  Õ«·… «·ÿ·»"
Else
MsgBox "Please Select Order Status"
End If
DcbStatusMaint.SetFocus
Exit Sub
End If

If val(DcbTypeMaint.ListIndex) = -1 Then
'If SystemOptions.UserInterface = ArabicInterface Then
'MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «·’Ì«‰…"
'Else
'MsgBox "Please Select Type Maintenance"
'End If
'DcbTypeMaint.SetFocus
'Exit Sub
End If
            my_branch = Me.dcBranch.BoundText
If val(DcbEquepment.BoundText) <> 0 Then
If CheckEqupRequest(val(DcbEquepment.BoundText)) = True Then
MsgBox "·«Ì„ﬂ‰ › Õ ÿ·» ’Ì«‰… ÃœÌœ ·Â–Â «·„⁄œ….ÌÊÃœ ÿ·» ”«»ﬁ „› ÊÕ «Ê  Õ  «· ‰›Ì– "
Exit Sub
End If
End If

If StopLocation.Text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· »Ì«‰«  „ﬂ«‰ «·⁄ÿ·"
StopLocation.SetFocus
Exit Sub
End If
If TxtDes.Text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· »Ì«‰«  «·⁄ÿ·"
TxtDes.SetFocus
Exit Sub
End If
If TxtDes.Text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· »Ì«‰«  «·⁄ÿ·"
TxtDes.SetFocus
Exit Sub
End If
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5
        
        Load FrmSearchRequerMainten
       FrmSearchRequerMainten.show

        Case 6
            Unload Me

        Case 7
           ' ShowGL_cc Me.txtNoteSerial.text, , 200

        Case 8
            
            
      VIEW_ATTACH
                 Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.XPTxtID.Text) <> 0 Then
                print_report val(Me.XPTxtID.Text)
        
        
            End If
        
    End Select

    Exit Sub
ErrTrap:
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

    MySQL = " SELECT   expectedEndDate,expectedEndtime, dbo.TblRequerMainten.Mobile , tripKM,StopLocation,RemainKmToArrive,LastKM,ManualKM,DifferentKm,UpdateKM ,  EquipmentStatusid, dbo.TblRequerMainten.ID, dbo.TblRequerMainten.ProblemTimID, dbo.TblRequerMainten.ProblemOther, dbo.TblRequerMainten.StopTime, "
    MySQL = MySQL & " dbo.TblRequerMainten.StartTime, dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate,"
    MySQL = MySQL & " dbo.TblRequerMainten.StopDate, dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
    MySQL = MySQL & " dbo.TblRequerMainten.EquepID, dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.TblRequerMainten.BranchID, TblBranchesData_1.branch_name,"
    MySQL = MySQL & " TblBranchesData_1.branch_namee, dbo.FixedAssets.namee, dbo.TblRequerMainten.ExternaExam, dbo.TblRequerMainten.RemarksEqup,"
    MySQL = MySQL & " dbo.TblRequerMainten.OutCounter, dbo.TblRequerMainten.EnterCounter, dbo.TblRequerMainten.RejecReason, dbo.TblRequerMainten.DrievName,"
    MySQL = MySQL & " dbo.TblRequerMainten.DrievType, dbo.TblRequerMainten.LeaderType, dbo.TblRequerMainten.LeaderName, dbo.TblRequerMainten.OperationNo,"
    MySQL = MySQL & " dbo.TblRequerMainten.StatusMaint, dbo.TblRequerMainten.TypeMaint, dbo.TblRequerMainten.BranchIDTo, TblBranchesData_1.branch_name AS branch_nameTo,"
    MySQL = MySQL & " TblBranchesData_1.branch_namee AS branch_nameToE, dbo.TblRequerMainten.OperiatorID, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode,"
    MySQL = MySQL & " TblEmployee_2.Emp_Namee, dbo.TblRequerMainten.LeaderID, TblEmployee_1.Emp_Name AS LeaderEmp_Name, TblEmployee_1.Fullcode AS LeaderFullcode,"
    MySQL = MySQL & " TblEmployee_1.Emp_Namee AS LeaderEmp_NameE, dbo.TblRequerMainten.DrievID, TblEmployee_2.Emp_Name AS DrivEmp_Name,"
    MySQL = MySQL & " TblEmployee_2.Fullcode AS DrivFullcode, TblEmployee_2.Emp_Namee AS DrivEmp_NameE, dbo.TblRequerMaintenDet.PartID, FixedAssets_1.Name AS EqupName,TblRequerMainten.NoOfLabs ,TblRequerMainten.supervisorNotes ,"
    MySQL = MySQL & " FixedAssets_1.code AS Equpcode, FixedAssets_1.namee AS EqupNameE"
    MySQL = MySQL & " FROM dbo.FixedAssets FixedAssets_1 INNER JOIN"
    MySQL = MySQL & " dbo.TblRequerMaintenDet ON FixedAssets_1.id = dbo.TblRequerMaintenDet.PartID RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblRequerMainten ON dbo.TblRequerMaintenDet.ReqID = dbo.TblRequerMainten.ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee TblEmployee_2 ON dbo.TblRequerMainten.DrievID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee TblEmployee_1 ON dbo.TblRequerMainten.LeaderID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee TblEmployee_3 ON dbo.TblRequerMainten.OperiatorID = TblEmployee_3.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblBranchesData TblBranchesData_1 ON dbo.TblRequerMainten.BranchIDTo = TblBranchesData_1.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblBranchesData TblBranchesData_2 ON dbo.TblRequerMainten.BranchID = TblBranchesData_2.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.FixedAssets ON dbo.TblRequerMainten.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID"
    MySQL = MySQL & " Where (dbo.TblRequerMainten.id =" & val(Me.XPTxtID.Text) & ") "


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepRequerMainten.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepRequerMainten.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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

Private Sub CmdHelp_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments XPTxtID, "0703201702"

End Sub

Private Sub DcbDrievID_Change()
DcbDrievID_Click (0)
End Sub

Private Sub DcbDrievID_Click(Area As Integer)
 If val(DcbDrievID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbDrievID.BoundText, EmpCode
    Text6.Text = EmpCode
End Sub

Private Sub DcbDrievID_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lblType = 39
        FrmEmployeeSearch.show
  
    End If
End Sub





Private Sub EndActTIme_Click()
ODERTimeFocus = True
End Sub



 

Private Sub LastKM_Change()
calcDiffernt
End Sub

Private Sub ManualKM_Change()
calcDiffernt
End Sub
Function calcDiffernt()
DifferentKm = val(ManualKM.Text) - val(LastKM.Text)
End Function


Private Sub NoOfLabs_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, NoOfLabs.Text, 0)
End Sub


Private Sub txtLetter1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select

End Sub
Sub EmptyTxt()
Me.txtNum1.Text = ""
Me.txtNum2.Text = ""
Me.txtNum3.Text = ""
Me.txtNum4.Text = ""
Me.txtLetter1.Text = ""
Me.txtLetter2.Text = ""
Me.txtLetter3.Text = ""
Me.txtLetter4.Text = ""
End Sub
Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
Cal_Board
End Sub

Private Sub txtLetter4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum1_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum2_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub

Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
Cal_Board
End Sub

Private Sub txtNum3_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board
End Sub
Private Sub Cal_Board()
    TxtBoardNO.Text = txtLetter1.Text & " " & txtLetter2.Text & " " & txtLetter3.Text & " " & txtLetter4.Text & " " & txtNum1.Text & " " & txtNum2.Text & " " & txtNum3.Text & " " & txtNum4.Text
    RetriveCarsInfo , , TxtBoardNO.Text, 2
End Sub
Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
Cal_Board

End Sub

Private Sub txtNum4_KeyUp(KeyCode As Integer, Shift As Integer)
Cal_Board

End Sub

Private Sub txtNum4_LostFocus()
TxtBoardNO.SetFocus
End Sub
Private Sub dcbDrievType_Click(Index As Integer)
If dcbDrievType(0).value = True Then
Text6.Enabled = True
DcbDrievID.Enabled = True
TxtDrievName.Enabled = False
TxtDrievName.Text = ""
ElseIf dcbDrievType(1).value = True Then
Text6.Enabled = False
DcbDrievID.Enabled = False
TxtDrievName.Enabled = True
DcbDrievID.BoundText = 0
Text6.Text = ""
End If
End Sub

Sub RetriveCarsInfo(Optional CarID As Double = 0, Optional OperNo As String, Optional BoardNO As String, Optional Typ As Integer = 0)
If Me.TxtModFlg <> "R" Then
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblCarsData"
If Typ = 0 Then
sql = sql & "  Where FixedassetId = " & CarID & ""
ElseIf Typ = 1 Then
sql = sql & " where OperatorN='" & OperNo & "'"
ElseIf Typ = 2 Then
sql = sql & " where BoardNO='" & BoardNO & "'"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then

Me.LastKM.Text = IIf(IsNull(Rs3("LastKMCounter").value), "", Rs3("LastKMCounter").value)

If Typ <> 1 Then
TxtOperationNo.Text = IIf(IsNull(Rs3("OperatorN").value), "", Rs3("OperatorN").value)
End If
If Typ <> 2 Then
TxtBoardNO.Text = IIf(IsNull(Rs3("BoardNO").value), "", Rs3("BoardNO").value)
End If
If Typ <> 0 Then
DcbEquepment.BoundText = IIf(IsNull(Rs3("FixedassetId").value), 0, Rs3("FixedassetId").value)
End If
DcbBranchIDTo.BoundText = IIf(IsNull(Rs3("Branch_NO").value), 0, Rs3("Branch_NO").value)
DcbLeaderID.BoundText = IIf(IsNull(Rs3("Emp_id").value), 0, Rs3("Emp_id").value)
Else
DcbLeaderID.BoundText = 0
DcbBranchIDTo.BoundText = 0
If Typ <> 1 Then
TxtOperationNo.Text = ""
End If
If Typ <> 2 Then
TxtBoardNO.Text = ""
End If
If Typ <> 0 Then
DcbEquepment.BoundText = 0
End If
End If
End If
End Sub

Public Sub DcbEquepment_Change()
    DcbEquepment_Click (0)
    Retrive_CarParts
End Sub
Function GetSumRad(Optional CarID As Double) As Double
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     SUM(CountOrders) AS SumOrder"
sql = sql & " From dbo.TblOrderUpload"
sql = sql & " Where (CarID = " & CarID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetSumRad = IIf(IsNull(rs2("SumOrder").value), 0, rs2("SumOrder").value)
Else
GetSumRad = 0
End If
End Function
Private Sub DcbEquepment_Click(Area As Integer)
If val(Me.DcbEquepment.BoundText) <> 0 Then
RetriveCarsInfo val(Me.DcbEquepment.BoundText), , 0
If Me.TxtModFlg.Text <> "R" Then
NoOfLabs.Text = GetSumRad(GetCarID())
End If
End If
End Sub
Function GetCarID() As Double
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " select ID from TblCarsData where fixedAssetid =" & val(DcbEquepment.BoundText) & ""
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetCarID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
Else
GetCarID = 0
End If
End Function
Private Sub DcbEquepment_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        Load FrmCasrShearches
        FrmCasrShearches.SendForm = "RequerMainten"
        FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcbLeaderID_Change()
DcbLeaderID_Click (0)
End Sub
Function GetMobile(Optional Emp_id As Double) As String
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select Emp_mobile from TblEmployee where Emp_ID=" & Emp_id & " "
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetMobile = IIf(IsNull(rs2("Emp_mobile").value), "", rs2("Emp_mobile").value)
Else
GetMobile = ""
End If
End Function
Private Sub DcbLeaderID_Click(Area As Integer)
If Me.TxtModFlg.Text <> "R" Then
 If val(DcbLeaderID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbLeaderID.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    Me.TxtMobile.Text = GetMobile(val(DcbLeaderID.BoundText))
 End If
End Sub

Private Sub DcbLeaderID_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lblType = 38
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub dcbLeaderType_Click(Index As Integer)
If dcbLeaderType(0).value = True Then
TxtSearchCode.Enabled = True
DcbLeaderID.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.Text = ""
ElseIf dcbLeaderType(1).value = True Then
TxtSearchCode.Enabled = False
DcbLeaderID.Enabled = False
TxtLeaderName.Enabled = True
DcbLeaderID.BoundText = 0
TxtSearchCode.Text = ""
End If
End Sub

Private Sub DcbOperiatorID_Change()
DcbOperiatorID_Click (0)
End Sub

Private Sub DcbOperiatorID_Click(Area As Integer)

 If val(DcbOperiatorID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbOperiatorID.BoundText, EmpCode
    Text3.Text = EmpCode
End Sub

Private Sub DcbOperiatorID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lblType = 37
        FrmEmployeeSearch.show
  
    End If
End Sub

Private Sub DcbStatusMaint_Change()
TxtRejecReason.Visible = False
lbl(21).Visible = False
If val(DcbStatusMaint.ListIndex) = 1 Then
TxtRejecReason.Visible = True
lbl(21).Visible = True
End If
End Sub

Private Sub DcbStatusMaint_Click()
DcbStatusMaint_Change
End Sub
Private Sub Form_Load()

    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    Dim My_SQL As String
    Dim GrdBack As ClsBackGroundPic

    On Error GoTo ErrTrap
    
    Set GrdBack = New ClsBackGroundPic
    If SystemOptions.CanChangeStatusDateRequest = True Then
    DcbStatusMaint.Enabled = True
    XPDtbTrans.Enabled = True
    ODERdATEFocus = True
    
    Else
    ODERdATEFocus = False
    XPDtbTrans.Enabled = False
    DcbStatusMaint.Enabled = False
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        With DcbStatusMaint
            .Clear
            .AddItem "„› ÊÕ"
            .AddItem "„—›Ê÷"
            .AddItem " „ «·«‰ Â«¡"
            .AddItem " Õ  «· ‰›Ì–"
        End With

        With DcbTypeMaint
            .Clear
            .AddItem "œ«Œ·Ì"
            .AddItem "Œ«—ÃÌ"
        End With
    Else
        With DcbStatusMaint
            .Clear
            .AddItem "Open"
            .AddItem "Rejected"
            .AddItem "Completed"
        End With

        With DcbTypeMaint
            .Clear
            .AddItem "Internal"
            .AddItem "External"
        End With
    End If

    'Frame6.Visible = False
    Set TTD = New clstooltipdemand
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set Dcombos = New ClsDataCombos
  
    Dcombos.GetEmpDepartments Me.DcbUnit
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetEquipments DcbEquepment
    Dcombos.GetUsers Me.DCboUserName
    'Dcombos.GetEmployees Me.DcbLeaderID
    'Dcombos.GetEmployees Me.DcbDrievID
    Dcombos.GetEmployees Me.DcbOperiatorID
    Dcombos.GetBranches Me.DcbBranchIDTo
    
    Dim str  As String
    
    If SystemOptions.UserInterface = ArabicInterface Then
        str = " SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
        str = str & " dbo.TblEmployee.Emp_Namee"
    Else
        str = " SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
        str = str & " dbo.TblEmployee.Emp_Name"
    End If
    str = str & " FROM dbo.TblEmployee LEFT OUTER JOIN"
    str = str & " dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
 
    If SystemOptions.ShowDriverOnly = True Then
        str = str & " where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
    End If
    
    fill_combo DcbDrievID, str
    fill_combo DcbLeaderID, str
    'Dcombos.GetFileCustomer Me.DcbCustomer
    
    If SystemOptions.usertype <> UserAdminAll Then
        Me.dcBranch.Enabled = True
    End If
    SetDtpickerDate Me.XPDtbTrans
    'YearMonth
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblRequerMainten     Order By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPDtbTrans.value = Date
    Me.TxtModFlg.Text = "R"
            
    If SystemOptions.UserInterface = EnglishInterface Then
        ProblemTimID.AddItem "During Production"
        ProblemTimID.AddItem "During Start up"
        ProblemTimID.AddItem "During Repair"
        ProblemTimID.AddItem "Others"
            
        EquipmentStatusid.AddItem "working"
        EquipmentStatusid.AddItem "Stopped"
            
        SetInterface Me
        ChangeLang
    Else
        ProblemTimID.AddItem "«À‰«¡ «· ’‰Ì⁄"
        ProblemTimID.AddItem "«À‰«¡ »œ¡ «· ‘€Ì·"
        ProblemTimID.AddItem "«À‰«¡ «·«’·«Õ"
        ProblemTimID.AddItem "«Œ—Ï"
        
        EquipmentStatusid.AddItem " ⁄„·"
        EquipmentStatusid.AddItem "„ Êﬁ›…"
        EquipmentStatusid.AddItem "⁄ÿ· ÿ—Ìﬁ"
        EquipmentStatusid.AddItem "Õ«œÀ "
        
    End If
    
 
    
  
    Retrive
    
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    
    Exit Sub

ErrTrap:
End Sub
Function CheckEqupRequest(Optional EquepID As Double) As Boolean
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
CheckEqupRequest = False
sql = " Select * from TblRequerMainten where EquepID=" & EquepID & " and ID<>" & val(XPTxtID.Text) & " "
sql = sql & "  and (StatusMaint=0 or StatusMaint=3 )"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckEqupRequest = True
Else
CheckEqupRequest = False
End If
End Function
Private Sub ChangeLang()
lbl(23).Caption = "Parts"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
lbl(10).Caption = "Meter Out"
lbl(13).Caption = "Meter In"
lbl(5).Caption = "Type"
lbl(24).Caption = "Mobile"
Frame5.Caption = "Date & Time Process"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    lbl(19).Caption = "Oper. No."
 Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"
    lbl(20).Caption = "Order Status"
    Me.Label2.Caption = "Observer"
  lbl(17).Caption = "External Testing"
  lbl(14).Caption = "Notes on Equipment"
   Me.Label1.Caption = "Branch Request Repair"
dcbLeaderType(0).RightToLeft = False
dcbLeaderType(1).RightToLeft = False
dcbLeaderType(0).Caption = "Employee"
dcbLeaderType(1).Caption = "Not Employee"
dcbDrievType(0).RightToLeft = False
dcbDrievType(1).RightToLeft = False
dcbDrievType(0).Caption = "Employee"
dcbDrievType(1).Caption = "Not Employee"
Me.Frame2.Caption = "Equepment Leader "
Frame4.Caption = "Conductor Equipment"
    Me.Caption = "Maintenance Request"
    EleHeader.Caption = Me.Caption
    lbl(4).Caption = "OPR#"
    lbl(1).Caption = "Date"
    lblBr.Caption = "Branch-Based Repair"
   lbl(25).Caption = " This Screen For Recording Maintenance Requests"
   lbl(0).Caption = "Department"
 '  lbl(10).Caption = "Customer"
   Frame3.Caption = "Data"
   lbl(3).Caption = "TimeProblem"
   lbl(29).Caption = "Machine"
   lbl(2).Caption = "Problem"
   lbl(18).Caption = "Technical Notes"
   Frame5.Caption = "Date and Time Problem"
   lbl(9).Caption = "DateWork"
    lbl(15).Caption = "TimeWork"
     lbl(12).Caption = "DateStop"
      lbl(16).Caption = "TimeStop"
  Cmd(8).Caption = "Attachments"
 

 lbl(8).Caption = "By"
  
        lbl(7).Caption = "Curr rec."
    lbl(6).Caption = "rec. count"
    
    '########################## khaled ##############################
    With VSFlexGrid13
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
    End With
    removeRow.Caption = "Delete"
    clearGridBtn.Caption = "Delete All"
    showAll.Caption = "Show All"
    lbl(22).Caption = " License plate"
End Sub
 
Private Sub Form_Paint()
    TTD.Destroy
End Sub

Private Sub Form_Resize()
    TTD.Destroy
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

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
Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text3.Text, EmpID
        DcbOperiatorID.BoundText = EmpID
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.Text, EmpID
        DcbDrievID.BoundText = EmpID
    End If
End Sub

Private Sub TxtBoardNO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , , TxtBoardNO.Text, 2
End If
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
        Frame1.Enabled = False
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  "
            'Me.menue(2).Enabled = True
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
          '  TxtAdvanceValue.Locked = True
            Me.DcboBox.locked = True
          '  XPDtbTrans.Enabled = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If

        Case "N"
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
           ' XPDtbTrans.Enabled = True
          '  XPDtbTrans.value = Date

        Case "E"
        Frame1.Enabled = True
            '        Me.Caption = "  «” »Ì«‰ ⁄‰ „ÊŸ›  (  ⁄œÌ· )"
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
          '  TxtAdvanceValue.Locked = False
            Me.DcboBox.locked = False
         '   XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Private Sub TxtOperationNo_KeyPress(KeyAscii As Integer)
Dim ID As Double
If Me.TxtModFlg.Text <> "R" Then
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , Me.TxtOperationNo.Text, , 1
End If
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcbLeaderID.BoundText = EmpID
    End If
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
      On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
     Dim RsDetails1 As ADODB.Recordset
   Dim ContactTime As Date
   Dim expectedEndtime As Date
   
    Dim I As Integer
    Dim StrSQL As String
      EmptyTxt

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If Lngid <> 0 Then
            rs.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
  If Not IsNull(rs("StartTime").value) Then
      ContactTime = FormatDateTime(rs("StartTime").value, vbShortTime)
       Me.StartTime.value = ContactTime
   
    End If
'     If Not IsNull(rs("EndExptedTime").value) Then
'      ContactTime = FormatDateTime(rs("EndExptedTime").value, vbShortTime)
'        Me.EndExptedTime.value = ContactTime
   
'    End If
  If Not IsNull(rs("StopTime").value) Then
      ContactTime = FormatDateTime(rs("StopTime").value, vbShortTime)
        Me.EndActTIme.value = ContactTime
    End If
    
    
    expectedEndDate.value = IIf(IsNull(rs("expectedEndDate").value), Date, rs("expectedEndDate").value)
    
 If Not IsNull(rs("expectedEndtime").value) Then
      expectedEndtime = FormatDateTime(rs("expectedEndtime").value, vbShortTime)
        Me.expectedEndtime.value = expectedEndtime
    End If
       
       
    
    XPTxtID.Text = IIf(IsNull(rs("Id").value), "", (rs("Id").value))
    TxtMobile.Text = IIf(IsNull(rs("Mobile").value), "", (rs("Mobile").value))
    XPDtbTrans.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
    StartDate.value = IIf(IsNull(rs("StartDate").value), Date, rs("StartDate").value)
    EndExptedDate.value = IIf(IsNull(rs("StopDate").value), Date, rs("StopDate").value)
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    DcbEquepment.BoundText = IIf(IsNull(rs("EquepID").value), "", rs("EquepID").value)
    DcbUnit.BoundText = IIf(IsNull(rs("UnitID").value), "", rs("UnitID").value)
    ProblemTimID.ListIndex = val(IIf(IsNull(rs("ProblemTimID").value), -1, rs("ProblemTimID").value))
    EquipmentStatusid.ListIndex = val(IIf(IsNull(rs("EquipmentStatusid").value), -1, rs("EquipmentStatusid").value))
    
    
    StopLocation.Text = IIf(IsNull(rs("StopLocation").value), "", (rs("StopLocation").value))
    RemainKmToArrive.Text = IIf(IsNull(rs("RemainKmToArrive").value), 0, (rs("RemainKmToArrive").value))
    LastKM.Text = IIf(IsNull(rs("LastKM").value), 0, (rs("LastKM").value))
    ManualKM.Text = IIf(IsNull(rs("ManualKM").value), 0, (rs("ManualKM").value))
    DifferentKm.Text = IIf(IsNull(rs("DifferentKm").value), 0, (rs("DifferentKm").value))
    tripKM.Text = IIf(IsNull(rs("tripKM").value), 0, (rs("tripKM").value))
    
    NoOfLabs.Text = IIf(IsNull(rs("NoOfLabs").value), 0, (rs("NoOfLabs").value))
    supervisorNotes.Text = IIf(IsNull(rs("supervisorNotes").value), "", (rs("supervisorNotes").value))
    
    If IsNull(rs("UpdateKM").value) Then
    UpdateKM.value = vbUnchecked
    
    Else
                If (rs("UpdateKM").value) = vbFalse Then
                    UpdateKM.value = vbUnchecked
             Else
                UpdateKM.value = vbChecked
                End If
    
      
    End If
    
    UpdateKM.value = 1
    
    
    
    
 '   Me.DcbCustomer.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.TxtProblemOther.Text = IIf(IsNull(rs("ProblemOther").value), "", rs("ProblemOther").value)
    Me.TxtDes.Text = IIf(IsNull(rs("Des").value), "", rs("Des").value)
    Me.TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
   '''/28 04 2016
  '  DcbTypeMaint.ListIndex = IIf(IsNull(rs("TypeMaint").value), -1, rs("TypeMaint").value)
    DcbBranchIDTo.BoundText = IIf(IsNull(rs("BranchIDTo").value), "", rs("BranchIDTo").value)
    DcbOperiatorID.BoundText = IIf(IsNull(rs("OperiatorID").value), "", rs("OperiatorID").value)
    DcbStatusMaint.ListIndex = IIf(IsNull(rs("StatusMaint").value), -1, rs("StatusMaint").value)
    TxtOperationNo.Text = IIf(IsNull(rs("OperationNo").value), "", rs("OperationNo").value)
    DcbLeaderID.BoundText = IIf(IsNull(rs("LeaderID").value), "", rs("LeaderID").value)
    TxtLeaderName.Text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
    DcbDrievID.BoundText = IIf(IsNull(rs("DrievID").value), "", rs("DrievID").value)
    TxtDrievName.Text = IIf(IsNull(rs("DrievName").value), "", rs("DrievName").value)
    TxtRejecReason.Text = IIf(IsNull(rs("RejecReason").value), "", rs("RejecReason").value)
    TxtEnterCounter.Text = IIf(IsNull(rs("EnterCounter").value), "", rs("EnterCounter").value)
    TxtOutCounter.Text = IIf(IsNull(rs("OutCounter").value), "", rs("OutCounter").value)
    TxtRemarksEqup.Text = IIf(IsNull(rs("RemarksEqup").value), "", rs("RemarksEqup").value)
    TxtExternaExam.Text = IIf(IsNull(rs("ExternaExam").value), "", rs("ExternaExam").value)
    TxtBoardNO.Text = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
    If Not (IsNull(rs("DrievType").value)) Then
    If rs("DrievType").value = 1 Then
    dcbDrievType(1).value = True
    Else
    dcbDrievType(0).value = True
    End If
    End If
    If Not (IsNull(rs("LeaderType").value)) Then
    If rs("LeaderType").value = 1 Then
    dcbLeaderType(1).value = True
    Else
    dcbLeaderType(0).value = True
    End If
    End If
   

'       If IsNull(rs("posted").value) Then
'                                                   If SystemOptions.UserInterface = ArabicInterface Then
'                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
'                                                  Else
'                                                    Accredit.Caption = " send to Approval   "
'                                               End If
'                                               Accredit.Enabled = True
'  Else
''                                                   If SystemOptions.UserInterface = ArabicInterface Then
 '                                                   Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
 '                                                 Else
 ''                                                   Accredit.Caption = " sent to Approval   "
  '                                             End If
  '                                             Accredit.Enabled = False
  ' End If
  '
  '
  '  Set RsDetails = New ADODB.Recordset
 'StrSQL = " SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblEmployee.Emp_Code, "
'StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Fullcode, dbo.TblRegDateDelgateDails.remark,"
'StrSQL = StrSQL & "                      dbo.TblRegDateDelgateDails.Type"
'StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
'StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblEmployee.Emp_ID"
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.Type = 0) And (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"


' RsDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    Fg.Clear flexClearScrollable, flexClearEverything
'    Fg.Rows = Fg.FixedRows
'
'    If Not (RsDetails.BOF Or RsDetails.EOF) Then
'        RsDetails.MoveFirst
'        Fg.Rows = Fg.FixedRows + RsDetails.RecordCount

'        For i = Me.Fg.FixedRows To Fg.Rows - 1
'        Fg.TextMatrix(i, Fg.ColIndex("Serial")) = i
'        Fg.TextMatrix(i, Fg.ColIndex("remarks")) = IIf(IsNull(RsDetails("remark").value), "", RsDetails("remark").value) ' RsDetails("remark").value
'            Fg.TextMatrix(i, Fg.ColIndex("code")) = IIf(IsNull(RsDetails("fullcode").value), "", RsDetails("fullcode").value) 'RsDetails("fullcode").value
'            If SystemOptions.UserInterface = EnglishInterface Then
'           Fg.TextMatrix(i, Fg.ColIndex("empname")) = IIf(IsNull(RsDetails("Emp_Namee").value), "", RsDetails("Emp_Namee").value) 'RsDetails("Emp_Namee").value
'           Else
'           Fg.TextMatrix(i, Fg.ColIndex("empname")) = IIf(IsNull(RsDetails("emp_name").value), "", RsDetails("emp_name").value) ' RsDetails("emp_name").value
'           End If
'            Fg.TextMatrix(i, Fg.ColIndex("empid")) = RsDetails("EmpID").value
'            RsDetails.MoveNext
'        Next i
'
'    End If

'    RsDetails.Close
'    Set RsDetails = Nothing
   '''''''''''''///////////////////////
'   Set RsDetails1 = New ADODB.Recordset
' StrSQL = "SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblRegDateDelgateDails.remark, "
'StrSQL = StrSQL & "                      dbo.TblRegDateDelgateDails.Type , dbo.TblCompo.name, dbo.TblCompo.namee, dbo.TblRegDateDelgateDails.Quantity"
'StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
'  StrSQL = StrSQL & "                    dbo.TblCompo ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblCompo.Id"
'
'StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.Type = 1) And (dbo.TblRegDateDelgateDails.DelgID = " & val(Me.XPTxtID.text) & " )"



' RsDetails1.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 
'    fg2.Clear flexClearScrollable, flexClearEverything
'    fg2.Rows = fg2.FixedRows
'
'    If Not (RsDetails1.BOF Or RsDetails1.EOF) Then
'        RsDetails1.MoveFirst
'        fg2.Rows = fg2.FixedRows + RsDetails1.RecordCount
'
'        For i = Me.fg2.FixedRows To fg2.Rows - 1
'        fg2.TextMatrix(i, fg2.ColIndex("Serial")) = i
'        fg2.TextMatrix(i, fg2.ColIndex("remarks")) = IIf(IsNull(RsDetails1("remark").value), "", RsDetails1("remark").value) ' RsDetails1("remark").value
'            fg2.TextMatrix(i, fg2.ColIndex("code")) = IIf(IsNull(RsDetails1("quantity").value), "", RsDetails1("quantity").value) 'RsDetails1("fullcode").value
'            If SystemOptions.UserInterface = EnglishInterface Then
'           fg2.TextMatrix(i, fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("namee").value), "", RsDetails1("namee").value) 'RsDetails1("Emp_Namee").value
''           Else
 '          fg2.TextMatrix(i, fg2.ColIndex("name")) = IIf(IsNull(RsDetails1("name").value), "", RsDetails1("name").value) ' RsDetails1("emp_name").value
 '          End If
 '           fg2.TextMatrix(i, fg2.ColIndex("empid")) = RsDetails1("EmpID").value
 '           RsDetails1.MoveNext
 '       Next i
'
'    End If

'    RsDetails1.Close
'    Set RsDetails1 = Nothing
 '  fillapprovData
    
    Dim rs_det As ADODB.Recordset
    Set rs_det = New ADODB.Recordset
    
    StrSQL = " SELECT     dbo.TblRequerMaintenDet.ID, dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.FixedAssets.namee, dbo.TblRequerMaintenDet.ReqID"
    StrSQL = StrSQL & "  FROM         dbo.TblRequerMaintenDet INNER JOIN"
    StrSQL = StrSQL & "                   dbo.TblRequerMainten ON dbo.TblRequerMaintenDet.ReqID = dbo.TblRequerMainten.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.TblRequerMaintenDet.PartID = dbo.FixedAssets.id"
    StrSQL = StrSQL & " where TblRequerMainten.ID = " & XPTxtID.Text & " "
    
    rs_det.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    VSFlexGrid13.Clear
    VSFlexGrid13.Rows = 1
    If rs_det.RecordCount > 0 Then
        rs_det.MoveFirst
        With VSFlexGrid13
            .Rows = rs_det.RecordCount + 1
            For I = 1 To .Rows - 1
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs_det("ID").value), 0, rs_det("ID").value)
               ' .TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(rs_det("PartID").value), 0, rs_det("PartID").value)
                .TextMatrix(I, .ColIndex("Code")) = IIf(IsNull(rs_det("code").value), "", rs_det("code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_det("Name").value), "", rs_det("Name").value)
                Else
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_det("namee").value), "", rs_det("namee").value)
                End If
                rs_det.MoveNext
            Next
         End With
    End If
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub
Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim RsDetails1 As ADODB.Recordset
    Dim I As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String

    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
        If Me.DcbEquepment.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «”„ «·„⁄œÂ..!! "
            Else
                Msg = "Equipment name must be selected"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DcbEquepment.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
        If Me.ProblemTimID.ListIndex = -1 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ Êﬁ  «·„‘ﬂ·… ..!! "
            Else
                Msg = "Problem Time must be specified"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            ProblemTimID.SetFocus
            'SendKeys "{F4}"
            Exit Sub
        End If
        Dim RsTest As New ADODB.Recordset

        Cn.BeginTrans
        BeginTrans = True
    
        If TxtModFlg.Text = "N" Then
            XPTxtID.Text = CStr(new_id("TblRequerMainten", "ID", "", True))
            'TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
            rs.AddNew
        End If
    
        rs("ID").value = val(XPTxtID.Text)
        rs("RecordDate").value = XPDtbTrans.value
        rs("StartDate").value = StartDate.value
        
        rs("StopDate").value = EndExptedDate.value
        rs("UserID").value = IIf(Me.DCboUserName.BoundText = "", Null, Me.DCboUserName.BoundText)
        rs("BranchID").value = IIf(Me.dcBranch.BoundText = "", Null, Me.dcBranch.BoundText)
        rs("EquepID").value = IIf(Me.DcbEquepment.BoundText = "", Null, Me.DcbEquepment.BoundText)
        rs("UnitID").value = IIf(Me.DcbUnit.BoundText = "", Null, Me.DcbUnit.BoundText)
        rs("StartTime").value = FormatDateTime(Me.StartTime.value, vbShortTime)
        rs("StopTime").value = FormatDateTime(Me.EndActTIme.value, vbShortTime)
        
        rs("expectedEndDate").value = expectedEndDate.value
        rs("expectedEndtime").value = FormatDateTime(Me.expectedEndtime.value, vbShortTime)
        
        rs("ProblemOther").value = IIf(Me.TxtProblemOther.Text = "", "", Me.TxtProblemOther.Text)
        rs("Mobile").value = IIf(Me.TxtMobile.Text = "", Null, Me.TxtMobile.Text)
        rs("Des").value = IIf(Me.TxtDes.Text = "", "", Me.TxtDes.Text)
        rs("Remarks").value = IIf(Me.TxtRemarks.Text = "", "", Me.TxtRemarks.Text)
        rs("ProblemTimID").value = val(IIf(Me.ProblemTimID.ListIndex = -1, -1, Me.ProblemTimID.ListIndex))
        rs("EquipmentStatusid").value = val(IIf(Me.EquipmentStatusid.ListIndex = -1, -1, Me.EquipmentStatusid.ListIndex))
        
        
        rs("StopLocation").value = IIf(Me.StopLocation.Text = "", "", Me.StopLocation.Text)
        rs("RemainKmToArrive").value = IIf(Me.RemainKmToArrive.Text = "", 0, val(Me.RemainKmToArrive.Text))
        rs("LastKM").value = IIf(Me.LastKM.Text = "", 0, val(Me.LastKM.Text))
        rs("ManualKM").value = IIf(Me.ManualKM.Text = "", 0, val(Me.ManualKM.Text))
        rs("DifferentKm").value = IIf(Me.DifferentKm.Text = "", 0, val(Me.DifferentKm.Text))
        rs("tripKM").value = IIf(Me.tripKM.Text = "", 0, val(Me.tripKM.Text))
        
        If UpdateKM.value = vbChecked Then
         rs("UpdateKM").value = 1
        Else
        rs("UpdateKM").value = 0
        End If
        
        '''/////28 04 2016
        'rs("TypeMaint").value = IIf(val(Me.DcbTypeMaint.ListIndex) = -1, Null, Me.DcbTypeMaint.ListIndex)
        rs("BranchIDTo").value = IIf(Me.DcbBranchIDTo.BoundText = "", Null, Me.DcbBranchIDTo.BoundText)
        rs("OperiatorID").value = IIf(Me.DcbOperiatorID.BoundText = "", Null, Me.DcbOperiatorID.BoundText)
        rs("StatusMaint").value = IIf(val(Me.DcbStatusMaint.ListIndex) = -1, Null, Me.DcbStatusMaint.ListIndex)
        rs("OperationNo").value = IIf(Me.TxtOperationNo.Text = "", Null, TxtOperationNo.Text)
        rs("LeaderID").value = IIf(Me.DcbLeaderID.BoundText = "", Null, Me.DcbLeaderID.BoundText)
        rs("LeaderName").value = IIf(Me.TxtLeaderName.Text = "", Null, Me.TxtLeaderName.Text)
        rs("DrievID").value = IIf(Me.DcbDrievID.BoundText = "", Null, Me.DcbDrievID.BoundText)
        rs("DrievName").value = IIf(Me.TxtDrievName.Text = "", Null, Me.TxtDrievName.Text)
        rs("RejecReason").value = IIf(Me.TxtRejecReason.Text = "", Null, Me.TxtRejecReason.Text)
        rs("EnterCounter").value = IIf(Me.TxtEnterCounter.Text = "", Null, Me.TxtEnterCounter.Text)
        rs("OutCounter").value = IIf(Me.TxtOutCounter.Text = "", Null, Me.TxtOutCounter.Text)
        rs("RemarksEqup").value = IIf(Me.TxtRemarksEqup.Text = "", Null, Me.TxtRemarksEqup.Text)
        rs("ExternaExam").value = IIf(Me.TxtExternaExam.Text = "", Null, Me.TxtExternaExam.Text)
        rs("BoardNO").value = IIf(Me.TxtBoardNO.Text = "", Null, Me.TxtBoardNO.Text)
        
        rs("supervisorNotes").value = IIf(Me.supervisorNotes.Text = "", Null, Me.supervisorNotes.Text)
        rs("NoOfLabs").value = IIf(Me.NoOfLabs.Text = "", 0, val(Me.NoOfLabs.Text))
        
        If dcbLeaderType(1).value = True Then
            rs("LeaderType").value = 1
        Else
            rs("LeaderType").value = 0
        End If
        If dcbDrievType(1).value = True Then
            rs("DrievType").value = 1
        Else
            rs("DrievType").value = 0
        End If
        rs.update
        
        '######################################################### khaled Det #################################################################
        Dim rs_det As ADODB.Recordset
        Set rs_det = New ADODB.Recordset
    
        If TxtModFlg.Text = "E" Then
             Cn.Execute "delete TblRequerMaintenDet where ReqID = " & val(Me.XPTxtID.Text)
        End If
    
        StrSQL = "SELECT  *  From TblRequerMaintenDet"
    
        rs_det.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        With VSFlexGrid13
            For I = 1 To .Rows - 1
                rs_det.AddNew
                rs_det("ReqID").value = IIf(Me.XPTxtID.Text = 0, Null, Me.XPTxtID.Text)
                rs_det("PartID").value = IIf(.TextMatrix(I, .ColIndex("PartID")) = "", Null, .TextMatrix(I, .ColIndex("PartID")))
                rs_det.update
            Next
        End With
    
        Cn.CommitTrans
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        
                If val(ManualKM.Text) <> 0 And UpdateKM.value = vbChecked Then
         Cn.Execute "Update  TblCarsData set LastKMCounter=" & val(ManualKM.Text) & " where fixedAssetid=" & val(DcbEquepment.BoundText) & ""
        End If
        
       
        
        
        Select Case Me.TxtModFlg.Text
            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " This is Record already Saved" & CHR(13)
                    Msg = Msg + "you want to enter another Record"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
        End Select
        TxtModFlg.Text = "R"
    End If
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
            Msg = "Can Not Save Data" & CHR(13)
            Msg = Msg + "I have been insert incorrect data " & CHR(13)
            Msg = Msg + "Make sure of the accuracy of the data and try again"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry...error douring save" & CHR(13)
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
    
    
    'Set RsDetails = New ADODB.Recordset
    'StrSQL = "SELECT     *  from dbo.TblRegDateDelgateDails Where (1 = -1)"
    'RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    
    '     For i = Me.Fg.FixedRows To Fg.Rows - 1
   '    If val(Fg.TextMatrix(i, Fg.ColIndex("EmpID"))) <> 0 Then
   '         RsDetails.AddNew
   '         RsDetails("DelgID").value = val(XPTxtID.text)
   '         RsDetails("Type").value = 0
   '        RsDetails("remark").value = Fg.TextMatrix(i, Fg.ColIndex("remarks"))
   '         RsDetails("EmpID").value = val(Fg.TextMatrix(i, Fg.ColIndex("empid")))
   '
   '         RsDetails.update
   '     End If
   '     Next i
  ''///////////'''''''''''''''''''''''''''''''
   '     Set RsDetails1 = New ADODB.Recordset
   ''    StrSQL = "SELECT     *  from dbo.TblRegDateDelgateDails Where (1 = -1)"
 '  RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
    
 '       For i = Me.fg2.FixedRows To fg2.Rows - 1
 '      If val(fg2.TextMatrix(i, fg2.ColIndex("EmpID"))) <> 0 Then
 '           RsDetails1.AddNew
 '           RsDetails1("DelgID").value = val(XPTxtID.text)
 '           RsDetails1("Type").value = 1
 '          RsDetails1("remark").value = fg2.TextMatrix(i, fg2.ColIndex("remarks"))
 '           RsDetails1("EmpID").value = val(fg2.TextMatrix(i, fg2.ColIndex("empid")))
 '   RsDetails1("quantity").value = val(fg2.TextMatrix(i, fg2.ColIndex("code")))
 '           RsDetails1.update
 '       End If
 '       Next i
'        Dim NoteID As Long
'        Dim line_no As Integer
'        Dim RsNotes As New ADODB.Recordset
'        RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
'        If detect_employee_work_type = 1 Then
        
'            If Me.TxtModFlg.text = "E" Then
 
'                StrSQL = "Delete notes where NoteID=" & val(Me.TxtNoteID.text)
'                Cn.Execute StrSQL, , adExecuteNoRecords

'            End If

'            RsNotes.AddNew
'            NoteID = CStr(TxtNoteID.text)
'            RsNotes("NoteID").value = CStr(TxtNoteID.text)
'            RsNotes("NoteType").value = 8032
'            RsNotes("NoteDate").value = XPDtbTrans.value
'            RsNotes("UserID").value = user_id
'            RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
'            RsNotes("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text) '„”·”· «–‰ «·’—›
'            RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'            RsNotes("numbering_type1").value = sand_numbering_type(32) ' ”ÃÌ· «·”·›'‰Ê⁄  —ﬁÌ„    
'            RsNotes("sanad_year").value = year(XPDtbTrans.value)
'            RsNotes("sanad_month").value = Month(XPDtbTrans.value)
'            RsNotes("note_value_by_characters").value = WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
            '     RsNotes("remark").value = txtRemarks.text & bankDes
'            RsNotes("Branch_no").value = val(Me.Dcbranch.BoundText)
                
'            RsNotes.update
                
'            line_no = 1
        
'            Msg = "”·› „ÊŸ›Ì‰ —ﬁ„ " & val(Me.XPTxtID.text)
'            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'            Employee_account = get_EMPLOYEE_Account(val(Me.DcboEmpName.BoundText), "Account_Code")
'            StrAccountCode = Employee_account
'
            '        StrAccountCode = "a1a3a4" 'Õ”«» “„„ «·„ÊŸ›Ì‰
'            If ModAccounts.AddNewDev(LngDevID, 1, StrAccountCode, val(Me.TxtAdvanceValue.text), 0, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If

'            StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

'            If ModAccounts.AddNewDev(LngDevID, 2, StrAccountCode, val(Me.TxtAdvanceValue.text), 1, Msg, NoteID, , , , Me.XPDtbTrans.value, Me.DCboUserName.BoundText, , , val(Me.XPTxtID.text), , , , , , , , , , , , , , , val(Me.Dcbranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
        
'        End If
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "ID='" & val(XPTxtID.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String
    Dim StrSQL1 As String
    
    On Error GoTo ErrTrap
If val(DcbStatusMaint.ListIndex) = 2 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«Ì„ﬂ‰ «·Õ–› .·ﬁœ  „ «·«‰ Â«¡ „‰ ÿ·» «·’Ì«‰Â"
Else
MsgBox "You can note delete.This is process completed"
End If
Exit Sub
End If

    If XPTxtID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Confirm Delete"
        End If
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From TblRequerMainten Where ID=" & val(Me.XPTxtID.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                'det part
                Cn.Execute "delete TblRequerMaintenDet where ReqID = " & val(Me.XPTxtID.Text)
                rs.MoveFirst
                If rs.RecordCount < 1 Then
                    clear_all Me
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
            Msg = "This process is not available There are no records"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If
    TxtModFlg_Change
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry error douring Delete data"
    End If
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub



'Function FillApprovedTable()
' Dim RSApproval  As New ADODB.Recordset
'   Set RSApproval = New ADODB.Recordset
'   Dim currentdate As Date
'   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'
' Dim sql As String
'  Dim rs1 As New ADODB.Recordset
' Dim i As Integer
'    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
'  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
'  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
'  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
'  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
'sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.name & "')"
'sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "
'
'    rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs1.RecordCount > 0 Then
'            currentdate = Now
'            For i = 1 To rs1.RecordCount
'              RSApproval.AddNew
'                RSApproval("ScreenName").value = Me.name
'                RSApproval("levelo").value = IIf(IsNull(rs1("levelo").value), Null, rs1("levelo").value)
'               RSApproval("EmpID").value = IIf(IsNull(rs1("EmpID").value), Null, rs1("EmpID").value)
'                RSApproval("levelorder").value = IIf(IsNull(rs1("levelorder").value), Null, rs1("levelorder").value)
'                 RSApproval("currorder").value = IIf(IsNull(rs1("currorder").value), Null, rs1("currorder").value)
'                  RSApproval("Transaction_ID").value = val(Me.XPTxtID.text)
'                   RSApproval("NoteSerial").value = val(Me.XPTxtID.text)
'                RSApproval("Transaction_Date").value = Date
                
'                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.name), currentdate)
'               RSApproval("SendTime").value = currentdate
'
'                 If i = 1 Then
'                        RSApproval("Currcursor").value = 1
'                         RSApproval("FromUser").value = user_name
'                End If
'
'                RSApproval.update
'                rs1.MoveNext
'            Next i
'
'    End If
    
    

'End Function



'Function fillapprovData()
'Dim Num As Integer
' Dim RsDetails As New ADODB.Recordset
' Dim StrSQL As String
'
'
' StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
'StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
'StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
'StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
'StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(Me.XPTxtID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
'StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
' If Not (RsDetails.EOF Or RsDetails.BOF) Then
'        GRID2.Rows = RsDetails.RecordCount + 1
'
'
'        For Num = 1 To RsDetails.RecordCount
'
'       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
'    If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
'   GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
'   Else
'    GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
'    End If
'
'        GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
'           If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
'          Else
'             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
'          End If
'            If SystemOptions.UserInterface = ArabicInterface Then
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            Else
'            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
'            End If
'            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
'          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
'
'
'RsDetails.MoveNext
'If Num = RsDetails.RecordCount Then
'
'        If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                      Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
'                                 Else
'                                       Label11.Caption = "Approved"
'                                 End If
'                            Label11.BackColor = &H80FF80
'        Else
'                             If SystemOptions.UserInterface = ArabicInterface Then
'                                     Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
'                            Else
'                                     Label11.Caption = "Currently required Approve"
'                            End If
'                 Label11.BackColor = &HFFFFC0
'        End If
'
'End If

'        Next Num
'Else
' GRID2.Rows = 1
'    End If
'RsDetails.Close
'
'End Function
'Private Sub ChekRepeat(Optional ind As Integer, Optional Row As Long, Optional ByRef bo As Boolean)
'    Dim i As Integer
'
'
'    With fg2
' bo = False
'        For i = .FixedRows To .Rows - 1
'If i <> Row Then
'            If val(.TextMatrix(i, .ColIndex("empid"))) = val(ind) Then
'             bo = True
'   End If
'            End If
'            Next i
'            End With
'        With Fg
' bo = False
'        For i = .FixedRows To .Rows - 1
'If i <> Row Then
'            If val(.TextMatrix(i, .ColIndex("empid"))) = val(ind) Then
'             bo = True
'             End If
'             Else
             
'            If val(ind) = val(Me.DcboEmpName.BoundText) Then
'              bo = True
'              End If
'   End If
'
'            Next i
'            End With
'        End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
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
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·⁄„·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "   ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap, True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "     ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "     ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "    ‘«‘… ÿ·» ’Ì«‰…   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
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
       
                Cmd_Click (2)

                ' btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub Retrive_CarParts()
    Dim I As Integer
    Set rs_CarParts = New ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = " SELECT     dbo.TblCarsDataDet.ID AS PID, dbo.TblCarsDataDet.PartID, dbo.FixedAssets.code, dbo.FixedAssets.Name, dbo.FixedAssets.namee"
    StrSQL = StrSQL & "   FROM         dbo.TblCarsDataDet LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.TblCarsDataDet.PartID = dbo.FixedAssets.id"
    StrSQL = StrSQL & " Where TblCarsDataDet.EqupID = " & val(Me.DcbEquepment.BoundText) & " "
    
    rs_CarParts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    VSFlexGrid13.Rows = 1
    If rs_CarParts.RecordCount > 0 Then
        rs_CarParts.MoveFirst
        With VSFlexGrid13
            .Rows = rs_CarParts.RecordCount + 1
            For I = 1 To .Rows - 1
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs_CarParts("PID").value), 0, rs_CarParts("PID").value)
                .TextMatrix(I, .ColIndex("PartID")) = IIf(IsNull(rs_CarParts("PartID").value), 0, rs_CarParts("PartID").value)
                .TextMatrix(I, .ColIndex("Code")) = IIf(IsNull(rs_CarParts("code").value), "", rs_CarParts("code").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("Name").value), "", rs_CarParts("Name").value)
                Else
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("namee").value), "", rs_CarParts("namee").value)
                End If
                rs_CarParts.MoveNext
            Next
         End With
    End If
End Sub
Private Sub showAll_Click()
    Retrive_CarParts
End Sub
Private Sub RemoveGridRow()
    With Me.VSFlexGrid13
        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With
End Sub
Private Sub cleargrid()
    VSFlexGrid13.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid13.Rows = 1
End Sub
Private Sub removeRow_Click()
    RemoveGridRow
End Sub
Private Sub clearGridBtn_Click()
    cleargrid
End Sub

Private Sub XPDtbTrans_Click()
ODERdATEFocus = True
End Sub
