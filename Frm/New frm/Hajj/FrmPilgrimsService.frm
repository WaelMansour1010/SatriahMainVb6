VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmPilgrimsService 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14235
   Icon            =   "FrmPilgrimsService.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9645
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
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Õ«›·… «·»œÌ·…"
      Height          =   1095
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   90
      Top             =   2880
      Width           =   5175
      Begin VB.TextBox TxtSwapOperaNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DcbSwapEqupID 
         Bindings        =   "FrmPilgrimsService.frx":6852
         Height          =   315
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
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
         Caption         =   "—ﬁ„ «·Õ«›·…"
         Height          =   285
         Index           =   21
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
         Height          =   285
         Index           =   20
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   600
         Width           =   1515
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmPilgrimsService.frx":6867
      Left            =   15480
      List            =   "FrmPilgrimsService.frx":6877
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   53
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
      TabIndex        =   47
      Top             =   0
      Width           =   14505
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
         TabIndex        =   48
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
         ButtonImage     =   "FrmPilgrimsService.frx":6890
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   49
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
         ButtonImage     =   "FrmPilgrimsService.frx":6C2A
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   50
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
         ButtonImage     =   "FrmPilgrimsService.frx":6FC4
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   51
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
         ButtonImage     =   "FrmPilgrimsService.frx":735E
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÃœÊ·  —ÕÌ· «·ÕÃ«Ã"
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
         TabIndex        =   52
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmPilgrimsService.frx":76F8
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   720
      Width           =   14235
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   4560
         Width           =   14055
         Begin VB.TextBox TxtLargNoPay 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   2640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtSamllNoPay 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   2640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   2355
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   13845
            _cx             =   24421
            _cy             =   4154
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPilgrimsService.frx":8AFD
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
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   4695
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   14055
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   5175
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   14055
            Begin VB.TextBox TxtSuperVisor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5640
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   240
               Width           =   6855
            End
            Begin VB.Frame Frame11 
               BackColor       =   &H00E2E9E9&
               Height          =   1455
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   2520
               Width           =   13935
               Begin VB.TextBox TxtNoPasPort 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   240
                  Width           =   1095
               End
               Begin VB.TextBox TxtCompNo 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   9600
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   1455
               End
               Begin VB.TextBox TxtOfficeNo 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   8640
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   960
                  Width           =   1455
               End
               Begin VB.TextBox TxtDepandNo 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   11040
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   960
                  Width           =   1455
               End
               Begin VB.TextBox TxtDcbEmploSearch 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   11430
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   240
                  Width           =   1065
               End
               Begin MSDataListLib.DataCombo DcbEmployee 
                  Bindings        =   "FrmPilgrimsService.frx":8BE5
                  Height          =   315
                  Left            =   5640
                  TabIndex        =   20
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
               Begin MSComCtl2.DTPicker TripDate 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   23
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   73465857
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal TripDateH 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   24
                  Top             =   600
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
               End
               Begin MSComCtl2.DTPicker StartTime 
                  Height          =   315
                  Left            =   2280
                  TabIndex        =   21
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   73465858
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker ByDate 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   28
                  Top             =   960
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   73465857
                  CurrentDate     =   38784
               End
               Begin Dynamic_Byte.NourHijriCal ByDateH 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   29
                  Top             =   960
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
               End
               Begin MSDataListLib.DataCombo DcbPath 
                  Bindings        =   "FrmPilgrimsService.frx":8BFA
                  Height          =   315
                  Left            =   5640
                  TabIndex        =   110
                  Top             =   600
                  Width           =   6855
                  _ExtentX        =   12091
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
               Begin MSDataListLib.DataCombo CompanyID 
                  Height          =   315
                  Left            =   5640
                  TabIndex        =   111
                  Top             =   960
                  Width           =   2250
                  _ExtentX        =   3969
                  _ExtentY        =   556
                  _Version        =   393216
                  Style           =   2
                  BackColor       =   16777215
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·ÃÊ«“« "
                  Height          =   285
                  Index           =   32
                  Left            =   1320
                  TabIndex        =   103
                  Top             =   240
                  Width           =   885
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·„ﬂ »"
                  Height          =   285
                  Index           =   31
                  Left            =   9960
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   960
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ƒ””…"
                  Height          =   285
                  Index           =   30
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   960
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·«⁄ „«œ"
                  Height          =   285
                  Index           =   29
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   960
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· Õ—Ì—"
                  Height          =   285
                  Index           =   28
                  Left            =   4050
                  TabIndex        =   99
                  Top             =   975
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Êﬁ  «·ﬁÌ«„"
                  Height          =   285
                  Index           =   27
                  Left            =   4050
                  TabIndex        =   98
                  Top             =   255
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·—Õ·…"
                  Height          =   285
                  Index           =   26
                  Left            =   4050
                  TabIndex        =   97
                  Top             =   615
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÃÂ… «·—Õ·… "
                  Height          =   285
                  Index           =   23
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   600
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ÊŸ› «·„Œ ’"
                  Height          =   285
                  Index           =   22
                  Left            =   12360
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   240
                  Width           =   1515
               End
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   11430
               RightToLeft     =   -1  'True
               TabIndex        =   2
               Top             =   240
               Visible         =   0   'False
               Width           =   1065
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ«›·…"
               Height          =   1335
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   120
               Width           =   5175
               Begin VB.TextBox TxtOperationNo 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Height          =   315
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   600
                  Width           =   3375
               End
               Begin MSDataListLib.DataCombo DcbEqupID 
                  Bindings        =   "FrmPilgrimsService.frx":8C0F
                  Height          =   315
                  Left            =   360
                  TabIndex        =   9
                  Top             =   240
                  Width           =   3375
                  _ExtentX        =   5953
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
               Begin MSDataListLib.DataCombo DcbPasseType 
                  Bindings        =   "FrmPilgrimsService.frx":8C24
                  Height          =   315
                  Left            =   360
                  TabIndex        =   11
                  Top             =   960
                  Width           =   3375
                  _ExtentX        =   5953
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
                  Caption         =   "‰Ê⁄ «·«—ﬂ«»"
                  Height          =   285
                  Index           =   13
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   960
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
                  Height          =   285
                  Index           =   1
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   88
                  Top             =   600
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ«›·…"
                  Height          =   285
                  Index           =   3
                  Left            =   3600
                  RightToLeft     =   -1  'True
                  TabIndex        =   87
                  Top             =   240
                  Width           =   1515
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁ«∆œ «·Õ«›·… «·»œÌ·"
               Height          =   1095
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   1560
               Width           =   8535
               Begin VB.TextBox TxtSwapName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   600
                  Width           =   6855
               End
               Begin VB.TextBox TxtSwapCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6030
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   240
                  Width           =   1065
               End
               Begin XtremeSuiteControls.RadioButton ChSwapType 
                  Height          =   255
                  Index           =   0
                  Left            =   6840
                  TabIndex        =   12
                  Top             =   240
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„ÊŸ›"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbSwapDrivID 
                  Bindings        =   "FrmPilgrimsService.frx":8C39
                  Height          =   315
                  Left            =   240
                  TabIndex        =   14
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
               Begin XtremeSuiteControls.RadioButton ChSwapType 
                  Height          =   255
                  Index           =   1
                  Left            =   6840
                  TabIndex        =   15
                  Top             =   600
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "€Ì— „ÊŸ›"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬁ«∆œ «·Õ«›·…"
               Height          =   975
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   600
               Width           =   8535
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6030
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   240
                  Width           =   1065
               End
               Begin VB.TextBox TxtDriverName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   600
                  Width           =   6855
               End
               Begin XtremeSuiteControls.RadioButton ChTypeDrive 
                  Height          =   255
                  Index           =   0
                  Left            =   6840
                  TabIndex        =   4
                  Top             =   240
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„ÊŸ›"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcbDriverID 
                  Bindings        =   "FrmPilgrimsService.frx":8C4E
                  Height          =   315
                  Left            =   240
                  TabIndex        =   6
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
               Begin XtremeSuiteControls.RadioButton ChTypeDrive 
                  Height          =   255
                  Index           =   1
                  Left            =   6840
                  TabIndex        =   7
                  Top             =   600
                  Width           =   1455
                  _Version        =   786432
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "€Ì— „ÊŸ›"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.TextBox TxtCodeUnit 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtDivArae 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtDevlopValue 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Enabled         =   0   'False
               Height          =   315
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   240
               Visible         =   0   'False
               Width           =   1455
            End
            Begin MSDataListLib.DataCombo DcbSuperViID 
               Bindings        =   "FrmPilgrimsService.frx":8C63
               Height          =   315
               Left            =   5640
               TabIndex        =   3
               Top             =   240
               Visible         =   0   'False
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
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„‰œÊ»"
               Height          =   285
               Index           =   16
               Left            =   12600
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·ÊÕœ…"
               Height          =   285
               Index           =   18
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   82
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
               TabIndex        =   80
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
               TabIndex        =   78
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
               TabIndex        =   77
               Top             =   240
               Visible         =   0   'False
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
               TabIndex        =   76
               Top             =   240
               Visible         =   0   'False
               Width           =   1515
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   120
         TabIndex        =   43
         Top             =   0
         Width           =   14055
         Begin VB.TextBox TxtSamllNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtLargNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   360
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtOrderNo 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   11640
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   9480
            TabIndex        =   0
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   73465857
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmPilgrimsService.frx":8C78
            Height          =   315
            Left            =   3000
            TabIndex        =   1
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
         Begin Dynamic_Byte.NourHijriCal RecordDateH 
            Height          =   315
            Left            =   7800
            TabIndex        =   83
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï «⁄ „«œ "
            Height          =   285
            Index           =   5
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   240
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   6480
            TabIndex        =   73
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
            TabIndex        =   45
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   10770
            TabIndex        =   44
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
      TabIndex        =   41
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
      TabIndex        =   40
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   55
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
      TabIndex        =   56
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
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   8280
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
         TabIndex        =   59
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   63
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
            TabIndex        =   62
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
            TabIndex        =   61
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
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
      End
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
            TabIndex        =   34
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
            ButtonImage     =   "FrmPilgrimsService.frx":8C8D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   8880
            TabIndex        =   36
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
            ButtonImage     =   "FrmPilgrimsService.frx":F4EF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   35
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
            ButtonImage     =   "FrmPilgrimsService.frx":F889
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7200
            TabIndex        =   37
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
            ButtonImage     =   "FrmPilgrimsService.frx":160EB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   5520
            TabIndex        =   38
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
            ButtonImage     =   "FrmPilgrimsService.frx":16485
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   39
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
            ButtonImage     =   "FrmPilgrimsService.frx":16A1F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   3960
            TabIndex        =   71
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
            ButtonImage     =   "FrmPilgrimsService.frx":16DB9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1920
            TabIndex        =   72
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
            ButtonImage     =   "FrmPilgrimsService.frx":1D61B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   9840
         TabIndex        =   64
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
         TabIndex        =   68
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ—— »Ê«”ÿ…  "
         Height          =   270
         Index           =   8
         Left            =   13080
         TabIndex        =   65
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
            Picture         =   "FrmPilgrimsService.frx":1D9B5
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPilgrimsService.frx":1DD4F
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPilgrimsService.frx":1E0E9
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPilgrimsService.frx":1E483
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPilgrimsService.frx":1E81D
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPilgrimsService.frx":1EBB7
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPilgrimsService.frx":1EF51
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPilgrimsService.frx":1F4EB
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   66
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
      ButtonImage     =   "FrmPilgrimsService.frx":1F885
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   69
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
      ButtonImage     =   "FrmPilgrimsService.frx":260E7
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   70
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
      ButtonImage     =   "FrmPilgrimsService.frx":2C949
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
      TabIndex        =   67
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmPilgrimsService"
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
 Dim RecID As String
 Dim Account_Code_dynamic As String
 Dim II As Long


Private Sub ByDate_Change()
If Me.TxtModFlg.text <> "R" Then
         ByDateH.value = ToHijriDate(ByDate.value)
End If
End Sub



Private Sub ByDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
      VBA.Calendar = vbCalGreg
    ByDate.value = ToGregorianDate(ByDateH.value)
    End If
End Sub

Private Sub ChSwapType_Click(Index As Integer)
TxtSwapCode.Enabled = False
DcbSwapDrivID.Enabled = False
TxtSwapName.Enabled = False
If ChSwapType(0).value = True Then
TxtSwapCode.Enabled = True
DcbSwapDrivID.Enabled = True
TxtSwapName.text = ""
Else
TxtSwapName.Enabled = True
TxtSwapCode.text = ""
DcbSwapDrivID.BoundText = 0
End If
End Sub

Private Sub ChTypeDrive_Click(Index As Integer)
TxtSearchCode.Enabled = False
DcbDriverID.Enabled = False
TxtDriverName.Enabled = False
If ChTypeDrive(0).value = True Then
DcbDriverID.Enabled = True
TxtSearchCode.Enabled = True
TxtDriverName.text = ""
Else
TxtDriverName.Enabled = True
DcbDriverID.BoundText = 0
TxtSearchCode.text = ""
End If
End Sub

Private Sub DcbDriverID_Change()
If val(DcbDriverID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbDriverID.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
End Sub
Private Sub DcbDriverID_Click(Area As Integer)
DcbDriverID_Change
End Sub

Private Sub DcbEmployee_Change()
If val(DcbEmployee.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbEmployee.BoundText, EmpCode
    TxtDcbEmploSearch.text = EmpCode
End Sub

Private Sub DcbEmployee_Click(Area As Integer)
DcbEmployee_Change
End Sub

Private Sub DcbEqupID_Change()
Dim fullcode As String
If val(Me.DcbEqupID.BoundText) <> 0 Then
RetriveCarsInfo val(Me.DcbEqupID.BoundText), fullcode
TxtOperationNo.text = fullcode
End If
End Sub

Private Sub DcbEqupID_Click(Area As Integer)
DcbEqupID_Change
End Sub

Private Sub DcbSuperViID_Change()
If val(DcbSuperViID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbSuperViID.BoundText, EmpCode
    Text3.text = EmpCode
End Sub

Private Sub DcbSuperViID_Click(Area As Integer)
DcbSuperViID_Change
End Sub

Private Sub DcbSwapDrivID_Change()
If val(DcbSwapDrivID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbSwapDrivID.BoundText, EmpCode
    TxtSwapCode.text = EmpCode
End Sub
Sub RetriveCarsInfo(Optional ByRef CarID As Double = 0, Optional ByRef OperNo As String, Optional typ As Integer = 0)
Dim Sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
If typ = 0 Then
Sql = "select OperatorN,ID from TblCarsData where ID=" & CarID & ""
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
OperNo = IIf(IsNull(Rs3("OperatorN").value), "", Rs3("OperatorN").value)
Else
OperNo = ""
End If
Else
Sql = "select OperatorN,id from TblCarsData where OperatorN='" & OperNo & "'"
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
CarID = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
Else
CarID = 0
End If

End If
End Sub

Private Sub DcbSwapEqupID_Change()
Dim fullcode As String
If val(Me.DcbSwapEqupID.BoundText) <> 0 Then
RetriveCarsInfo val(Me.DcbSwapEqupID.BoundText), fullcode
TxtSwapOperaNo.text = fullcode
End If
End Sub

Private Sub DcbSwapEqupID_Click(Area As Integer)
DcbSwapEqupID_Change
End Sub

    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
   If SystemOptions.UserInterface = ArabicInterface Then
                GridInstallments.ColComboList(GridInstallments.ColIndex("TypePilgrims")) = "#1;  ﬂ»«—|#2; ’€«—"
   ElseIf SystemOptions.UserInterface = EnglishInterface Then
               GridInstallments.ColComboList(GridInstallments.ColIndex("TypePilgrims")) = "#1;Large |#2; Small "
             
   End If
     Dim str  As String
      str = "  select   e.Emp_ID Emp_ID , e.Emp_Name   Emp_Name  from TblEmployee e, TblEmpJobsTypes  j"
      str = str & "   Where e.JobTypeID = j.JobTypeID"
      str = str & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
    fill_combo DcbDriverID, str
    fill_combo DcbSwapDrivID, str
         str = "  select   id, BoardNO from TblCarsData"
     fill_combo DcbSwapEqupID, str
     
    str = "  select   id, BoardNO from TblCarsData"
     fill_combo DcbEqupID, str
    If SystemOptions.UserInterface = ArabicInterface Then
    str = " select  ID , Name   from TblTourismCompanies "
    Else
    str = " Select ID , NameE from  TblTourismCompanies "
   End If
fill_combo CompanyID, str
    conection = "select * from TblPilgrimsService order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.text = "R"
    Resize_Form Me

    Dim Dcombos As New ClsDataCombos
    Dcombos.GetTblShrines Me.DcbPath
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.DcbEmployee
    Dcombos.GetEmployees Me.DcbSuperViID
    Dcombos.GetVehicleType Me.DcbPasseType
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
    Dim Sql As String
    Dim ID As Double
             If Me.TxtModFlg.text = "E" Then
                 StrSQL = "Delete From TblPilgrimsServiceDet Where PilgrSerID =" & val(TxtSerial1.text) & ""
                  Cn.Execute StrSQL, , adExecuteNoRecords
              End If
  
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("RecordDateH").value = RecordDateH.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    '''////
    RsSavRec.Fields("CompID").value = val(Me.CompanyID.BoundText)
    RsSavRec.Fields("SuperViID").value = val(DcbSuperViID.BoundText)
    RsSavRec.Fields("DriverID").value = val(Me.DcbDriverID.BoundText)
    RsSavRec.Fields("DriverName").value = TxtDriverName.text
    RsSavRec.Fields("SwapDrivID").value = val(Me.DcbSwapDrivID.BoundText)
    RsSavRec.Fields("SwapName").value = (TxtSwapName.text)
    RsSavRec.Fields("SuperVisor").value = TxtSuperVisor.text
 ''''//////////////////////
    RsSavRec.Fields("EqupID").value = val(DcbEqupID.BoundText)
    RsSavRec.Fields("OperationNo").value = (TxtOperationNo.text)
    RsSavRec.Fields("SwapOperaNo").value = (TxtSwapOperaNo.text)
    RsSavRec.Fields("SwapEqupID").value = val(Me.DcbSwapEqupID.BoundText)
    RsSavRec.Fields("PasseType").value = val(Me.DcbPasseType.BoundText)
    RsSavRec.Fields("EmpID").value = val(Me.DcbEmployee.BoundText)
 '   RsSavRec.Fields("FromCity").value = val(DcbFromCity.BoundText)
 '   RsSavRec.Fields("ToCity").value = val(DcbToCity.BoundText)
    RsSavRec.Fields("TripDate").value = TripDate.value
    RsSavRec.Fields("TripDateH").value = TripDateH.value
    RsSavRec.Fields("ByDate").value = ByDate.value
    RsSavRec.Fields("ByDateH").value = ByDateH.value
    RsSavRec.Fields("NoPasPort").value = val((Me.TxtNoPasPort.text))
    RsSavRec.Fields("OfficeNo").value = ((Me.TxtOfficeNo.text))
    RsSavRec.Fields("CompNo").value = ((Me.TxtCompNo.text))
    RsSavRec.Fields("DepandNo").value = ((Me.TxtDepandNo.text))
    RsSavRec.Fields("StartTime").value = FormatDateTime(Me.StartTime.value, vbShortTime)
If ChTypeDrive(1).value = True Then
RsSavRec.Fields("TypeDrive").value = 1
Else
RsSavRec.Fields("TypeDrive").value = 0
End If
If ChSwapType(1).value = True Then
RsSavRec.Fields("SwapType").value = 1
Else
RsSavRec.Fields("SwapType").value = 0
End If
RsSavRec.Fields("PathID").value = val(DcbPath.BoundText)
RsSavRec.Fields("OrderNo").value = val(TxtOrderNo.text)
RsSavRec.Fields("SamllNoPay").value = val(TxtSamllNoPay.text)
RsSavRec.Fields("LargNoPay").value = val(TxtLargNoPay.text)
RsSavRec.Fields("LargNo").value = val(TxtLargNo.text)
RsSavRec.Fields("SamllNo").value = val(TxtSamllNo.text)

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ''/////
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.Update

''//////////////////////////
Dim StrRecID As String
      Set RsDevsub = New ADODB.Recordset
    StrSQL = "SELECT  *  from TblPilgrimsServiceDet Where (1 = -1)"
    RsDevsub.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    Dim I As Integer
    Dim str2 As String
    With Me.GridInstallments
       For I = .FixedRows To .Rows - 1
       If val(.TextMatrix(I, .ColIndex("TypePilgrims"))) <> 0 Then
       StrRecID = new_id("TblPilgrimsServiceDet", "ID", "")
       RsDevsub.AddNew
       RsDevsub("id").value = StrRecID
                RsDevsub("PilgrSerID").value = val(TxtSerial1.text)
                RsDevsub("TypePilgrims").value = IIf((.TextMatrix(I, .ColIndex("TypePilgrims"))) = "", Null, val(.TextMatrix(I, .ColIndex("TypePilgrims"))))
                RsDevsub("NoPilgrims").value = IIf((.TextMatrix(I, .ColIndex("NoPilgrims"))) = "", Null, val(.TextMatrix(I, .ColIndex("NoPilgrims"))))
                RsDevsub("NationalID").value = IIf((.TextMatrix(I, .ColIndex("NationalID"))) = "", Null, val(.TextMatrix(I, .ColIndex("NationalID"))))
                RsDevsub("Remarks").value = IIf((.TextMatrix(I, .ColIndex("Remarks"))) = "", Null, (.TextMatrix(I, .ColIndex("Remarks"))))
       RsDevsub.Update
      End If
     Next I
    End With
'''///////////////
  
      Select Case Me.TxtModFlg.text
        Case "N"
            Dim Msg As String
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ï"
            Else
               Msg = " This record alredy saved... " & Chr(13)
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
    Dim I As Integer
    Dim ContactTime As Date
    TxtSerial1.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    Me.DcbSuperViID.BoundText = IIf(IsNull(RsSavRec.Fields("SuperViID").value), "", RsSavRec.Fields("SuperViID").value)
    Me.DcbDriverID.BoundText = IIf(IsNull(RsSavRec.Fields("DriverID").value), "", RsSavRec.Fields("DriverID").value)
    TxtDriverName.text = IIf(IsNull(RsSavRec.Fields("DriverName").value), "", RsSavRec.Fields("DriverName").value)
    DcbSwapDrivID.BoundText = IIf(IsNull(RsSavRec.Fields("SwapDrivID").value), 0, RsSavRec.Fields("SwapDrivID").value)
    TxtSwapName.text = IIf(IsNull(RsSavRec.Fields("SwapName").value), "", RsSavRec.Fields("SwapName").value)
    DcbEqupID.BoundText = IIf(IsNull(RsSavRec.Fields("EqupID").value), 0, RsSavRec.Fields("EqupID").value)
    TxtOperationNo.text = IIf(IsNull(RsSavRec.Fields("OperationNo").value), "", RsSavRec.Fields("OperationNo").value)
    TxtSwapOperaNo.text = IIf(IsNull(RsSavRec.Fields("SwapOperaNo").value), "", RsSavRec.Fields("SwapOperaNo").value)
    DcbSwapEqupID.BoundText = IIf(IsNull(RsSavRec.Fields("SwapEqupID").value), 0, RsSavRec.Fields("SwapEqupID").value)
    DcbPasseType.BoundText = IIf(IsNull(RsSavRec.Fields("PasseType").value), 0, RsSavRec.Fields("PasseType").value)
    Me.DcbEmployee.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), 0, RsSavRec.Fields("EmpID").value)
    Me.CompanyID.BoundText = IIf(IsNull(RsSavRec.Fields("CompID").value), 0, RsSavRec.Fields("CompID").value)
   ' Me.DcbFromCity.BoundText = IIf(IsNull(RsSavRec.Fields("FromCity").value), 0, RsSavRec.Fields("FromCity").value)
   ' Me.DcbToCity.BoundText = IIf(IsNull(RsSavRec.Fields("ToCity").value), 0, RsSavRec.Fields("ToCity").value)
    TripDate.value = IIf(IsNull(RsSavRec.Fields("TripDate").value), Date, RsSavRec.Fields("TripDate").value)
    TripDateH.value = IIf(IsNull(RsSavRec.Fields("TripDateH").value), ToHijriDate(Date), RsSavRec.Fields("TripDateH").value)
    ByDate.value = IIf(IsNull(RsSavRec.Fields("ByDate").value), Date, RsSavRec.Fields("ByDate").value)
    ByDateH.value = IIf(IsNull(RsSavRec.Fields("ByDateH").value), ToHijriDate(Date), RsSavRec.Fields("ByDateH").value)
    Me.TxtNoPasPort.text = IIf(IsNull(RsSavRec.Fields("NoPasPort").value), 0, RsSavRec.Fields("NoPasPort").value)
    TxtOfficeNo.text = IIf(IsNull(RsSavRec.Fields("OfficeNo").value), "", RsSavRec.Fields("OfficeNo").value)
    TxtCompNo.text = IIf(IsNull(RsSavRec.Fields("CompNo").value), "", RsSavRec.Fields("CompNo").value)
    TxtDepandNo.text = IIf(IsNull(RsSavRec.Fields("DepandNo").value), "", RsSavRec.Fields("DepandNo").value)
    TxtSuperVisor.text = IIf(IsNull(RsSavRec.Fields("SuperVisor").value), "", RsSavRec.Fields("SuperVisor").value)
     If Not IsNull(RsSavRec.Fields("StartTime").value) Then
      ContactTime = FormatDateTime(RsSavRec.Fields("StartTime").value, vbShortTime)
      Me.StartTime.value = ContactTime
    End If
    If Not (IsNull(RsSavRec.Fields("TypeDrive").value)) Then
    If RsSavRec.Fields("TypeDrive").value = 1 Then
    ChTypeDrive(1).value = True
    Else
    ChTypeDrive(0).value = True
    End If
    Else
    ChTypeDrive(0).value = True
    End If
    If Not (IsNull(RsSavRec.Fields("SwapType").value)) Then
    If RsSavRec.Fields("SwapType").value = 1 Then
    ChSwapType(1).value = True
    Else
    ChSwapType(0).value = True
    End If
    Else
    ChSwapType(0).value = True
    End If
   TxtSamllNo.text = IIf(IsNull(RsSavRec.Fields("SamllNo").value), 0, RsSavRec.Fields("SamllNo").value)
   TxtLargNo.text = IIf(IsNull(RsSavRec.Fields("LargNo").value), 0, RsSavRec.Fields("LargNo").value)
   TxtOrderNo.text = IIf(IsNull(RsSavRec.Fields("OrderNo").value), 0, RsSavRec.Fields("OrderNo").value)
   DcbPath.BoundText = IIf(IsNull(RsSavRec.Fields("PathID").value), 0, RsSavRec.Fields("PathID").value)
   TxtSamllNoPay.text = IIf(IsNull(RsSavRec.Fields("SamllNoPay").value), 0, RsSavRec.Fields("SamllNoPay").value)
   TxtLargNoPay.text = IIf(IsNull(RsSavRec.Fields("LargNoPay").value), 0, RsSavRec.Fields("LargNoPay").value)
    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60
FullGridData

ErrTrap:
End Sub
Function GetSmalNo(Optional ID As Double) As Double
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     OrderNo, SUM(SamllNoPay) AS SumLargNo"
Sql = Sql & " From dbo.TblPilgrimsService"
Sql = Sql & " Where (orderNo = " & ID & ") And (ID <> " & val(TxtSerial1.text) & ")"
Sql = Sql & " GROUP BY OrderNo"
Rs6.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
GetSmalNo = IIf(IsNull(Rs6("SumLargNo").value), 0, Rs6("SumLargNo").value)
Else
GetSmalNo = 0
End If
End Function
Function GetLargNo(Optional ID As Double) As Double
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Dim Sql As String
Sql = " SELECT     OrderNo, SUM(LargNoPay) AS SumLargNo"
Sql = Sql & " From dbo.TblPilgrimsService"
Sql = Sql & " Where (orderNo = " & ID & ") And (ID <> " & val(TxtSerial1.text) & ")"
Sql = Sql & " GROUP BY OrderNo"
Rs6.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
GetLargNo = IIf(IsNull(Rs6("SumLargNo").value), 0, Rs6("SumLargNo").value)
Else
GetLargNo = 0
End If
End Function
Private Sub ReLineGrid()
    Dim I As Integer
    Dim IntCounter  As Integer
    Dim LargNoPay As Integer
    Dim SamllNoPay As Integer
    IntCounter = 0
    LargNoPay = 0
    SamllNoPay = 0
    With GridInstallments

        For I = .FixedRows To .Rows - 1

            If val(.TextMatrix(I, .ColIndex("TypePilgrims"))) <> 0 Then
                IntCounter = IntCounter + 1
                .TextMatrix(I, .ColIndex("Ser")) = IntCounter
           If val(.TextMatrix(I, .ColIndex("TypePilgrims"))) = 1 Then
            LargNoPay = LargNoPay + val(.TextMatrix(I, .ColIndex("NoPilgrims")))
            If Round(LargNoPay, 2) > Round(val(TxtLargNo.text), 2) Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·⁄œœ «ﬂ»— „‰ «·⁄œœ ›Ì «·«⁄ „«œ"
            Else
            MsgBox "The No Larger than Toatl"
            End If
            .TextMatrix(I, .ColIndex("NoPilgrims")) = 0
            Exit Sub
            End If
            ElseIf val(.TextMatrix(I, .ColIndex("TypePilgrims"))) = 2 Then
            SamllNoPay = SamllNoPay + val(.TextMatrix(I, .ColIndex("NoPilgrims")))
                 If Round(SamllNoPay, 2) > Round(val(TxtSamllNo.text), 2) Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ «·⁄œœ «ﬂ»— „‰ «·⁄œœ ›Ì «·«⁄ „«œ"
            Else
            MsgBox "The No Larger than Toatl"
            End If
            .TextMatrix(I, .ColIndex("NoPilgrims")) = 0
            Exit Sub
            End If
            End If
            End If
        Next I
 
    End With
    TxtSamllNoPay.text = SamllNoPay
      TxtLargNoPay.text = LargNoPay

End Sub


Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim LngRow As Long
Dim StrAccountCode As String
With GridInstallments
Select Case .ColKey(Col)
Case "Name"
   StrAccountCode = .ComboData
                  LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("NationalID"), False, True)
                  .TextMatrix(Row, .ColIndex("NationalID")) = StrAccountCode
End Select
ReLineGrid
End With
End Sub

Private Sub GridInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With GridInstallments
        Select Case .ColKey(Col)
Case "NoPilgrims"
.ComboList = ""
Case "Remarks"
.ComboList = ""
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
Case "Name"
  StrSQL = "select * from Nationality "
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
              If Row = .Rows - 1 Then
                     .Rows = .Rows + 1
                     End If
    End With

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

    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            Dcbranch.SetFocus
            Exit Sub
     End If
           If TxtSuperVisor.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„  «·„‰œÊ»", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Eneter  Supervisor ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            TxtSuperVisor.SetFocus
            Exit Sub
     End If
    
       If DcbEmployee.text = "" And val(DcbEmployee.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·„ÊŸ› «·„Œ ’", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbEmployee.SetFocus
            Exit Sub
     End If
      If DcbPath.text = "" And val(DcbPath.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ÃÂ… «·—Õ·… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Flight destination ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbPath.SetFocus
            Exit Sub
     End If
    'If DcbToCity.text = "" And val(DcbToCity.BoundText) = 0 Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ÃÂ… «·—Õ·… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        Else
    '        MsgBox "Please Select Flight destination ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '     End If
    '        DcbToCity.SetFocus
    ''        Exit Sub
     'End If
    If val(TxtNoPasPort.text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡   «œŒ«· ⁄œœ «·ÃÊ«“«  ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Eneter Number of passport ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            TxtNoPasPort.SetFocus
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
    StrRecID = new_id("TblPilgrimsService", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub


 Sub FullGridData()
 On Error GoTo ErrTrap
  Dim Rs1 As ADODB.Recordset
  Set Rs1 = New ADODB.Recordset
  Dim Sql As String
    GridInstallments.Clear flexClearScrollable, flexClearEverything
    GridInstallments.Rows = 1
Sql = "SELECT     dbo.TblPilgrimsServiceDet.ID, dbo.TblPilgrimsServiceDet.PilgrSerID, dbo.TblPilgrimsServiceDet.TypePilgrims, dbo.TblPilgrimsServiceDet.NoPilgrims, "
Sql = Sql & "                      dbo.TblPilgrimsServiceDet.Remarks , dbo.TblPilgrimsServiceDet.NationalID, dbo.Nationality.name, dbo.Nationality.NameE"
Sql = Sql & " FROM         dbo.TblPilgrimsServiceDet LEFT OUTER JOIN"
Sql = Sql & "                      dbo.Nationality ON dbo.TblPilgrimsServiceDet.NationalID = dbo.Nationality.id"
Sql = Sql & " Where (dbo.TblPilgrimsServiceDet.PilgrSerID =" & val(TxtSerial1.text) & ") "

  Rs1.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Rs1.RecordCount > 0 Then
     Rs1.MoveFirst
     End If
     Dim I As Integer
     With Me.GridInstallments
                    For I = .FixedRows To Rs1.RecordCount
                   .Rows = .FixedRows + Rs1.RecordCount
                   .TextMatrix(I, .ColIndex("Ser")) = I
                   .TextMatrix(I, .ColIndex("TypePilgrims")) = IIf(IsNull(Rs1("TypePilgrims").value), 0, Rs1("TypePilgrims").value)
                   .TextMatrix(I, .ColIndex("NoPilgrims")) = IIf(IsNull(Rs1("NoPilgrims").value), 0, Rs1("NoPilgrims").value)
                   .TextMatrix(I, .ColIndex("Remarks")) = IIf(IsNull(Rs1("Remarks").value), "", Rs1("Remarks").value)
                   .TextMatrix(I, .ColIndex("NationalID")) = IIf(IsNull(Rs1("NationalID").value), 0, Rs1("NationalID").value)
                   If SystemOptions.UserInterface = ArabicInterface Then
                   .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs1("Name").value), "", Rs1("Name").value)
                   Else
                   .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(Rs1("NameE").value), "", Rs1("NameE").value)
                   End If
                   Rs1.MoveNext
             Next I
        End With
        Exit Sub
ErrTrap:
    End Sub



Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
      VBA.Calendar = vbCalGreg
    XPDtbTrans.value = ToGregorianDate(RecordDateH.value)
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text3.text, EmpID
        DcbSuperViID.BoundText = EmpID
    End If
End Sub



Private Sub TripDate_Change()
If Me.TxtModFlg.text <> "R" Then
         TripDateH.value = ToHijriDate(TripDate.value)
End If
End Sub

Private Sub TripDateH_LostFocus()
If Me.TxtModFlg.text <> "R" Then
      VBA.Calendar = vbCalGreg
    TripDate.value = ToGregorianDate(TripDateH.value)
    End If
End Sub

Private Sub TxtDcbEmploSearch_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtDcbEmploSearch.text, EmpID
        DcbEmployee.BoundText = EmpID
    End If
End Sub
Private Sub TxtOperationNo_KeyPress(KeyAscii As Integer)
Dim ID As Double
RetriveCarsInfo ID, TxtOperationNo.text, 1
DcbEqupID.BoundText = ID
End Sub





Private Sub TxtOrderNo_Change()
Dim str As String
If Me.TxtModFlg.text <> "R" Then
If val(TxtOrderNo.text) <> 0 Then
RetriveOrder val(TxtOrderNo.text)
TxtSamllNo.text = val(TxtSamllNo.text) - GetSmalNo(val(TxtOrderNo.text))
TxtLargNo.text = val(TxtLargNo.text) - GetLargNo(val(TxtOrderNo.text))
End If
End If
   str = " select   id, BoardNO from TblCarsData where id in(select CarID from TblEndorseTransDet where EnTransID=" & val(TxtOrderNo.text) & ")"
     fill_combo DcbEqupID, str
End Sub
Sub RetriveOrder(Optional ID As Double)
Dim Sql As String
Dim Rs6 As ADODB.Recordset
Set Rs6 = New ADODB.Recordset
Sql = "Select * from TblEndorseTrans where ID= " & ID & ""
Rs6.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs6.RecordCount > 0 Then
Me.CompanyID.BoundText = IIf(IsNull(Rs6("CompanyID").value), 0, Rs6("CompanyID").value)
TxtLargNo.text = IIf(IsNull(Rs6("TotOlds").value), 0, Rs6("TotOlds").value)
TxtSamllNo.text = IIf(IsNull(Rs6("TotYoungs").value), 0, Rs6("TotYoungs").value)
DcbPath.BoundText = IIf(IsNull(Rs6("PathID").value), 0, Rs6("PathID").value)
DcbPasseType.BoundText = IIf(IsNull(Rs6("VehicleType").value), 0, Rs6("VehicleType").value)
Else
DcbPasseType.BoundText = 0
CompanyID.BoundText = 0
TxtLargNo.text = 0
TxtSamllNo.text = 0
DcbPath.BoundText = 0
End If
End Sub


Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcbDriverID.BoundText = EmpID
    End If
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
    RsSavRec.find "ID=" & RecID, , adSearchForward, 1
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
    Dim Sql As String
    Dim Msg As String
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If
    Dim X As Integer
    Dim I As Integer
    Dim ID As Double

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If
    If X = vbNo Then Exit Sub
     If TxtSerial1.text = "" Then
       If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Nothing To Delet ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
               Else
                X = MsgBox("⁄›Ê« ...·« ÌÊÃœ »Ì«‰«  ··Õ–›", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
       End If
               Else
              
                StrSQL = "delete From TblPilgrimsServiceDet  where  PilgrSerID =" & val(TxtSerial1.text) & ""
                   Cn.Execute StrSQL, , adExecuteNoRecords
                RsSavRec.find "ID=" & val(TxtSerial1.text), , adSearchForward, 1
       
                                          RsSavRec.Delete
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
            LabCurrRec.Caption = 0
            LabCountRec.Caption = 0
                 If SystemOptions.UserInterface = EnglishInterface Then
                X = MsgBox(" Delete  Successfully ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
                Else
                X = MsgBox(" „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title)
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
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
                   RecID As String)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
            Msg = Msg & "Data will be updated"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
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
            GridInstallments.Rows = GridInstallments.Rows + 1
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
        Me.Dcbranch.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            'Could not update; currently locked.
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
            Msg = "Sorry.." & Chr(13)
            Msg = Msg & " You can not edit this the record now" & Chr(13)
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
    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    Dcbranch.BoundText = Current_branch
     ChTypeDrive(0).value = True
 ChSwapType(0).value = True
 ChTypeDrive_Click (0)
ChSwapType_Click (0)
    TxtModFlg.text = "N"
    GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 2
    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
  
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
               Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
            Msg = "Sorry I've been to delete this record" & Chr(13)
            Msg = Msg & "By another user on the network " & Chr(13)
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
  MySQL = " SELECT     dbo.TblPilgrimsService.ID, dbo.TblPilgrimsService.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
  MySQL = MySQL & "                    dbo.TblPilgrimsService.RecordDate, dbo.TblPilgrimsService.RecordDateH, dbo.TblPilgrimsService.SuperViID, dbo.TblEmployee.Emp_Name AS SuperEmp_Name,"
  MySQL = MySQL & "                    dbo.TblEmployee.Fullcode AS SuperFullcode, dbo.TblEmployee.Emp_Namee AS SuperEmp_NameE, dbo.TblPilgrimsService.TypeDrive,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.DriverName, dbo.TblPilgrimsService.DriverID, TblEmployee_1.Emp_Name AS DrivEmp_Name, TblEmployee_1.Fullcode AS DrivFullcode,"
  MySQL = MySQL & "                    TblEmployee_1.Emp_Namee AS DrivEmp_NameE, dbo.TblPilgrimsService.SwapType, dbo.TblPilgrimsService.SwapName, dbo.TblPilgrimsService.SwapDrivID,"
  MySQL = MySQL & "                    TblEmployee_2.Emp_Name AS SwapEmp_Name, TblEmployee_2.Fullcode AS SwapFullcode, TblEmployee_2.Emp_Namee AS SwapEmp_NameE,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.EqupID, dbo.TblCarsData.Branch_NO, dbo.TblCarsData.Fullcode AS CarFullcode, dbo.TblCarsData.BoardNO, dbo.TblCarsData.OperatorN,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.OperationNo, dbo.TblPilgrimsService.SwapOperaNo, dbo.TblPilgrimsService.SwapEqupID, TblCarsData_1.Fullcode AS SwCarFullcode,"
  MySQL = MySQL & "                    TblCarsData_1.BoardNO AS SwCarBoardNO, TblCarsData_1.OperatorN AS SwCarOperatorN, dbo.TblPilgrimsService.PasseType, dbo.TblVehicleType.Name,"
  MySQL = MySQL & "                    dbo.TblVehicleType.NameE, dbo.TblPilgrimsService.EmpID, TblEmployee_3.Emp_Name, TblEmployee_3.Fullcode, TblEmployee_3.Emp_Namee,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.TripDate, dbo.TblPilgrimsService.TripDateH, dbo.TblPilgrimsService.ByDate, dbo.TblPilgrimsService.ByDateH,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.NoPasPort, dbo.TblPilgrimsService.OfficeNo, dbo.TblPilgrimsService.DepandNo, dbo.TblPilgrimsService.StartTime,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.SamllNo, dbo.TblPilgrimsService.LargNo, dbo.TblPilgrimsService.OrderNo, dbo.TblPilgrimsService.SamllNoPay,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.LargNoPay, dbo.TblPilgrimsService.PathID, dbo.TblShrines.Name AS ShirName, dbo.TblShrines.NameE AS ShirNameE,"
  MySQL = MySQL & "                    dbo.TblPilgrimsService.CompID, dbo.TblTourismCompanies.Name AS ComoName, dbo.TblTourismCompanies.NameE AS ComoNameE,"
  MySQL = MySQL & "                    dbo.TblPilgrimsServiceDet.NoPilgrims, dbo.TblPilgrimsServiceDet.Remarks, dbo.TblPilgrimsServiceDet.TypePilgrims, dbo.TblPilgrimsServiceDet.NationalID,"
  MySQL = MySQL & "                    dbo.Nationality.name AS NationalName, dbo.Nationality.namee AS NationalNameE"
  MySQL = MySQL & "    FROM         dbo.Nationality RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblPilgrimsServiceDet ON dbo.Nationality.id = dbo.TblPilgrimsServiceDet.NationalID RIGHT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblPilgrimsService ON dbo.TblPilgrimsServiceDet.PilgrSerID = dbo.TblPilgrimsService.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblTourismCompanies ON dbo.TblPilgrimsService.CompID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblShrines ON dbo.TblPilgrimsService.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee TblEmployee_3 ON dbo.TblPilgrimsService.EmpID = TblEmployee_3.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblVehicleType ON dbo.TblPilgrimsService.PasseType = dbo.TblVehicleType.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCarsData TblCarsData_1 ON dbo.TblPilgrimsService.SwapEqupID = TblCarsData_1.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCarsData ON dbo.TblPilgrimsService.EqupID = dbo.TblCarsData.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee TblEmployee_2 ON dbo.TblPilgrimsService.SwapDrivID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee TblEmployee_1 ON dbo.TblPilgrimsService.DriverID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee ON dbo.TblPilgrimsService.SuperViID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblPilgrimsService.BranchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "  Where (dbo.TblPilgrimsService.id =" & val(TxtSerial1.text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPilgrimsService.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPilgrimsService.rpt"
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


Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
    Me.Caption = "Hajj Service"
    Me.lbl(4).Caption = "No"
    Me.lbl(2).Caption = "Date"
    Me.lbl(7).Caption = "Branch"
    Me.Label1(2).Caption = Me.Caption
    lbl(16).Caption = "SuperVisor"
    Frame9.Caption = "Vehicles"
lbl(3).Caption = "Vehicle.No"
lbl(1).Caption = "Op.No"
lbl(13).Caption = "Type"
Frame10.Caption = "Alternative Vehicles"
lbl(21).Caption = "Vehicle.No"
lbl(20).Caption = "Op.No"
Frame7.Caption = "Leader Vehicles"
ChTypeDrive(0).Caption = "Employee"
ChTypeDrive(1).Caption = "Non"
ChTypeDrive(0).RightToLeft = False
ChTypeDrive(1).RightToLeft = False
Frame8.Caption = "Alternative Leader Vehicles"
ChSwapType(0).Caption = "Employee"
ChSwapType(1).Caption = "Non"
ChSwapType(0).RightToLeft = False
ChSwapType(1).RightToLeft = False
lbl(27).Caption = "Time"
lbl(32).Caption = "Pas.No."
lbl(22).Caption = "Employee"
lbl(23).Caption = "Trip"
lbl(24).Caption = "From"
lbl(25).Caption = "To"
lbl(29).Caption = "Dependence No."
lbl(31).Caption = "Office No."
lbl(30).Caption = "Company"
lbl(26).Caption = "Trip Date"
lbl(28).Caption = "Date"
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
  .TextMatrix(0, .ColIndex("TypePilgrims")) = "Type Hajj "
  .TextMatrix(0, .ColIndex("NoPilgrims")) = "Hajj No."
  .TextMatrix(0, .ColIndex("Name")) = "Nationality"
  .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
  End With
ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblPilgrimsService"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.text = rs.RecordCount + 1
    Else
        TxtSerial1.text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ en

Private Sub TxtSwapCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSwapCode.text, EmpID
        DcbSwapDrivID.BoundText = EmpID
    End If
End Sub

Private Sub TxtSwapOperaNo_KeyPress(KeyAscii As Integer)
Dim ID As Double
RetriveCarsInfo ID, TxtSwapOperaNo.text, 1
DcbSwapEqupID.BoundText = ID
End Sub

Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.text <> "R" Then
         RecordDateH.value = ToHijriDate(XPDtbTrans.value)
End If
End Sub
