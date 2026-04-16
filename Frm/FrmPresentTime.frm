VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPresentTime 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„Ê«⁄Ìœ «·Õ÷Ê— Ê«·√‰’—«› "
   ClientHeight    =   6450
   ClientLeft      =   2685
   ClientTop       =   2475
   ClientWidth     =   9435
   HelpContextID   =   550
   Icon            =   "FrmPresentTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   9435
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„⁄·Ê„« "
      ForeColor       =   &H000000C0&
      Height          =   1965
      Index           =   2
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   4440
      Width           =   3285
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   1665
         Left            =   90
         TabIndex        =   56
         Top             =   210
         Width           =   3075
         _cx             =   5424
         _cy             =   2937
         Appearance      =   2
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPresentTime.frx":038A
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
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  «·⁄„·Ì…"
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
      Height          =   645
      Index           =   1
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   1770
      Width           =   5985
      Begin VB.TextBox TxtCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3960
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   240
         Width           =   1200
      End
      Begin MSComCtl2.DTPicker DTDate 
         Height          =   330
         Left            =   930
         TabIndex        =   1
         Top             =   240
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   582
         _Version        =   393216
         Format          =   100073475
         CurrentDate     =   38887
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„”·”·"
         Height          =   195
         Index           =   1
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «· ”ÃÌ·"
         Height          =   285
         Index           =   0
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.Frame FrmBrngTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ”ÃÌ· «·Õ÷Ê— ·„ÊŸ›"
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
      Height          =   2985
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   2460
      Width           =   6000
      Begin VB.CheckBox Chk 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ”ÃÌ· ≈‰’—«› «·„ÊŸ›"
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
         Height          =   225
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
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
         Height          =   1155
         Index           =   3
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1200
         Width           =   2880
         Begin MSComCtl2.DTPicker DtpDeparture 
            Height          =   330
            Left            =   660
            TabIndex        =   7
            Top             =   300
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            Format          =   100073475
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DtpDepHour 
            Height          =   405
            Left            =   660
            TabIndex        =   8
            Top             =   660
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            _Version        =   393216
            CustomFormat    =   "'Time: 'hh:mm tt"
            Format          =   100073475
            UpDown          =   -1  'True
            CurrentDate     =   39240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”«⁄…"
            Height          =   195
            Index           =   6
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   780
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÌÊ„"
            Height          =   225
            Index           =   2
            Left            =   2490
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   360
            Width           =   315
         End
      End
      Begin VB.TextBox TxtEmp_Code 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   3990
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "⁄›Ê« Ì—Ã∆ ≈œŒ«· ﬂÊœ «·„ÊŸ›"
         Top             =   270
         Width           =   630
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  Õ÷Ê— «·„ÊŸ›"
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
         Height          =   1155
         Index           =   4
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1200
         Width           =   2880
         Begin MSComCtl2.DTPicker DtpPresent 
            Height          =   330
            Left            =   510
            TabIndex        =   4
            Top             =   330
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            Format          =   100073475
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker DtpPresentHour 
            Height          =   405
            Left            =   510
            TabIndex        =   5
            Top             =   690
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            _Version        =   393216
            CustomFormat    =   "'Time: 'hh:mm tt"
            Format          =   100073475
            UpDown          =   -1  'True
            CurrentDate     =   39240
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”«⁄…"
            Height          =   195
            Index           =   9
            Left            =   2310
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   780
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÌÊ„"
            Height          =   225
            Index           =   1
            Left            =   2460
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   360
            Width           =   315
         End
      End
      Begin MSDataListLib.DataCombo DCEmp_Name 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo CboShift 
         Height          =   315
         Left            =   3000
         TabIndex        =   71
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   "DCEmp_Name"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   90
         TabIndex        =   69
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·≈‰’—«›:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Êﬁ  «·Õ÷Ê—:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2400
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3000
         TabIndex        =   66
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3000
         TabIndex        =   65
         Top             =   2670
         Width           =   1305
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· √ŒÌ— ⁄‰ «·⁄„·:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   2670
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ ”«⁄«  «·⁄„·:-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   2670
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00400040&
         Height          =   255
         Index           =   10
         Left            =   90
         TabIndex        =   62
         Top             =   870
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ì⁄«œ ≈‰’—«› «·„ÊŸ›:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         ForeColor       =   &H00400040&
         Height          =   255
         Index           =   8
         Left            =   90
         TabIndex        =   60
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ì⁄«œ Õ÷Ê— «·„ÊŸ›:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   600
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·œÊ«„ «Ê «·‘Ì› "
         Height          =   225
         Index           =   5
         Left            =   4710
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   53
         Top             =   2670
         Width           =   1395
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·„ÊŸ›"
         Height          =   195
         Index           =   3
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   255
         Index           =   0
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ê«⁄Ìœ «·Õ÷Ê— ›Ï «·‘—ﬂ…"
      Enabled         =   0   'False
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
      Height          =   1005
      Index           =   0
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   720
      Width           =   6000
      Begin VB.ComboBox CmbTime 
         Height          =   315
         ItemData        =   "FrmPresentTime.frx":03D7
         Left            =   2865
         List            =   "FrmPresentTime.frx":03D9
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   480
         Width           =   690
      End
      Begin VB.TextBox TxtMinute 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3585
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   480
      End
      Begin VB.TextBox TxtHour 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4110
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   480
      End
      Begin VB.ComboBox CmbTimeType 
         ForeColor       =   &H000000FF&
         Height          =   315
         ItemData        =   "FrmPresentTime.frx":03DB
         Left            =   1590
         List            =   "FrmPresentTime.frx":03DD
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÊﬁÌ  « ·Õ÷Ê—"
         Height          =   195
         Index           =   0
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label LabDayName 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Height          =   330
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   195
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "› —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "”"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "œ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   225
         Width           =   90
      End
      Begin VB.Label LabWork 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   105
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   390
         Width           =   1680
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   -15
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   -15
      Width           =   9525
      Begin VB.TextBox TxtMoveTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   150
         Width           =   765
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   315
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Text            =   "modflag"
         Top             =   30
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.TextBox TxtPresentTime_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   285
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Visible         =   0   'False
         Width           =   945
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   4815
         Top             =   105
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPresentTime.frx":03DF
               Key             =   "Emp_Name"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPresentTime.frx":0779
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPresentTime.frx":0B13
               Key             =   "Emp_Code"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPresentTime.frx":10AD
               Key             =   "Emp_Salary"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   345
         Left            =   105
         TabIndex        =   17
         Top             =   150
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmPresentTime.frx":1447
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   345
         Left            =   660
         TabIndex        =   16
         Top             =   150
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmPresentTime.frx":17E1
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   345
         Left            =   1980
         TabIndex        =   15
         Top             =   150
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmPresentTime.frx":1B7B
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   345
         Left            =   2565
         TabIndex        =   14
         Top             =   150
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   609
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   14871017
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
         ButtonImage     =   "FrmPresentTime.frx":1F15
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê«⁄Ìœ «·Õ÷Ê— Ê«·√‰’—«› "
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
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   3975
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   930
      Left            =   3420
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5490
      Width           =   6000
      _cx             =   10583
      _cy             =   1640
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
      Begin ImpulseButton.ISButton btnNew 
         Height          =   420
         Left            =   5250
         TabIndex        =   9
         Top             =   465
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   741
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
         ButtonImage     =   "FrmPresentTime.frx":22AF
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   420
         Left            =   3705
         TabIndex        =   0
         Top             =   465
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   741
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
         ButtonImage     =   "FrmPresentTime.frx":2649
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   420
         Left            =   4425
         TabIndex        =   10
         Top             =   465
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   741
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
         ButtonImage     =   "FrmPresentTime.frx":29E3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   420
         Left            =   2850
         TabIndex        =   11
         Top             =   465
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   741
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
         ButtonImage     =   "FrmPresentTime.frx":2D7D
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   420
         Left            =   2130
         TabIndex        =   12
         Top             =   465
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   741
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
         ButtonImage     =   "FrmPresentTime.frx":3117
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   420
         Left            =   1230
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   450
         Width           =   870
         _ExtentX        =   1535
         _ExtentY        =   741
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
         ButtonImage     =   "FrmPresentTime.frx":36B1
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   420
         Left            =   105
         TabIndex        =   13
         Top             =   465
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
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
         ButtonImage     =   "FrmPresentTime.frx":3A4B
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin MSDataListLib.DataCombo DCUser 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   3210
         TabIndex        =   51
         Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
         Top             =   30
         Width           =   1770
         _ExtentX        =   3122
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ—— »Ê«”ÿ…"
         Height          =   270
         Index           =   13
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   60
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   210
         Index           =   2
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   75
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "/"
         Height          =   210
         Index           =   1
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   75
         Width           =   165
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   60
         Width           =   495
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   75
         Width           =   450
      End
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      DataField       =   "project_no"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   120
      TabIndex        =   72
      Top             =   1200
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "account_no"
      BoundColumn     =   "Fullcode"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·„‘—Ê⁄"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1920
      TabIndex        =   73
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label LblZoneSetting 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   4080
      Width           =   3345
   End
End
Attribute VB_Name = "FrmPresentTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim cSearch As clsDCboSearch
Dim RecID As String
Dim II As Long
Dim TTP As New clstooltip

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String

    If DoPremis(Do_Delete, Me.name, True) = False Then
        Exit Sub
    End If

    If TxtPresentTime_ID.text <> "" Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbQuestion + vbMsgBoxRight, App.Title)
        
        If MSGType = vbYes Then
            RsSavRec.find "Present_ID=" & val(TxtPresentTime_ID.text), , adSearchForward, 1
            RsSavRec.delete
                     
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbInformation + vbMsgBoxRight, App.Title
            '------------------------------ Move Next ---------------------------.
            
            BtnNext_Click
        End If
    
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbExclamation + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtPresentTime_ID.text)
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
        FindRec val(TxtPresentTime_ID.text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If DoPremis(Do_Edit, Me.name, True) = False Then
        Exit Sub
    End If

    If TxtPresentTime_ID.text <> "" Then
        TxtModFlg = "E"
        Me.DCUser.BoundText = user_id
        Me.DCEmp_Name.SetFocus
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
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    clear_all Me
    
    Me.DtpDeparture.value = Date
    Me.DtpPresent.value = Date
    Me.DtpPresentHour.value = Time

    If IsDate(Me.lbl(8).Caption) = True Then
        Me.DtpPresentHour.value = Me.lbl(8).Caption
    End If

    Me.DtpDepHour.value = Time

    If IsDate(Me.lbl(10).Caption) = True Then
        Me.DtpDepHour.value = Me.lbl(10).Caption
    End If

    TxtModFlg.text = "N"
    Me.DCUser.BoundText = user_id
    My_SQL = "select * From tblPresentTime where Present_Type=0"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        TXTCode.text = rs.RecordCount + 1
    Else
        TXTCode.text = 1
    End If

    rs.Close
    DTDate_Click
    
    If TxtEmp_Code.Enabled = True Then
        TxtEmp_Code.SetFocus
    End If

ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtPresentTime_ID.text)
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
        FindRec val(TxtPresentTime_ID.text)
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
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnQuery_Click()
    FrmPresentTimeSearch.Show
End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String

    'Dim StrVacCode As String
    'Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Msg As String

    '---------------------- check if data Vaclete -----------------------
    If CmbTimeType.ListIndex = 0 Then

        For Each CtrlTxt In Me.Controls

            If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
                If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                    MsgBox CtrlTxt.Tag, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    CtrlTxt.SetFocus
                    Exit Sub
                End If
            End If

        Next
    
        If DCEmp_Name.BoundText = "" Then
            Msg = "⁄›Ê« Ì—ÃÏ  ÕœÌœ «”„ «·„ÊŸ›"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCEmp_Name.SetFocus
            Exit Sub
        End If

        If Me.CboShift.BoundText = "" Then
            Msg = "ÌÃ»  ÕœÌœ «·‘Ì›  «Ê «·œ«Ê„...."
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            CboShift.SetFocus
            Exit Sub
        End If

        If ChkEmpComeToday = False Then
            Msg = "⁄›Ê« Â–« «·„ÊŸ› ·„ ÌÕ÷— «·ÌÊ„"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        If ChkEmpExist = True Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· «·Õ÷Ê— ·Â–« «·„ÊŸ› " & Chr(13)
            Msg = Msg & "›Ï ‰›”  «·ÌÊ„ „‰ ﬁ»·"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        If Me.Chk.value = vbChecked Then
            If DateDiff("n", GetPresentTime, GetDepTime) < 0 Then
                Msg = "ÌÃ» «‰ ÌﬂÊ‰ Êﬁ  «·√‰’—«› »⁄œ Êﬁ  «·Õ÷Ê—...!"
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
        End If

        '-------------------------------------- txtmodflg type -------------------
        Select Case Me.TxtModFlg.text

                '------------------------------ new record ----------------------------
            Case "N"
                '------------------------- save record -----------------------------
                AddNewRec
                BtnLast_Click

            Case "E"
    
                '----------------------------- save edit -------------------------------
                FiLLRec
        End Select

    Else
        Msg = "⁄›Ê« Â–« ÌÊ„ ⁄ÿ·…"
        MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtPresentTime_ID.text)
    Me.TxtModFlg.text = "R"
End Sub

Private Sub Chk_Click()
    Me.DtpDeparture.Enabled = CBool(Me.Chk.value)
    Me.lbl(2).Enabled = CBool(Me.Chk.value)
    Me.lbl(6).Enabled = CBool(Me.Chk.value)
    DtpDepHour.Enabled = CBool(Me.Chk.value)
    Me.lbl(4).Caption = IIf(CBool(Me.Chk.value) = True, Me.lbl(4).Caption, "")
    CalculateTimes
End Sub

Private Sub DCEmp_Name_Change()
    GetEmpSetting
    CalculateTimes
End Sub

Private Sub DCEmp_Name_Click(Area As Integer)
    GetEmpSetting
End Sub

Private Sub DTDate_Change()
    DTDate_Click
End Sub

Private Sub DTDate_Click()
    LabDayName.Caption = Format(DtDate.value, "dddd")
    GetTimeDetails
End Sub

Private Sub DtpDeparture_Change()

    If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
        CalculateTimes
    End If

End Sub

Private Sub DtpDepHour_Change()

    If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
        CalculateTimes
    End If

End Sub

Private Sub DtpPresent_Change()

    If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
        CalculateTimes
    End If

End Sub

Private Sub DtpPresentHour_Change()

    If Me.TxtModFlg.text = "E" Or Me.TxtModFlg.text = "N" Then
        CalculateTimes
    End If

End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Attendance Registeration"
    Label1(2).Caption = Me.Caption

    LblZoneSetting.Visible = False
    Fra(2).Visible = False
    Fra(0).Caption = "attendance at the company"
    lbl(0).Caption = "Time"
    Label4(1).Caption = "H"
    Label4(0).Caption = "M"
    Label4(2).Caption = ""
    LabWork.Visible = False
    Fra(1).Caption = "OPR Details"
    Label1(1).Caption = "Ser"
    Label2(0).Caption = "Date"
    FrmBrngTime.Caption = "Register"
    Label1(3).Caption = "Emp Code"
    Label1(0).Caption = "Name"
    lbl(5).Caption = "Shift"
    lbl(3).Caption = "Presence"
    lbl(7).Caption = "Departure"
    Fra(4).Caption = " Presence Time"
    Chk.Caption = "Departure Time"
    lbl(1).Caption = "Day"
    lbl(9).Caption = "Time"
    lbl(2).Caption = "Day"
    lbl(6).Caption = "Time"

    lbl(14).Caption = " Presence"
    lbl(11).Caption = "Late"

    lbl(16).Caption = " Departure"
    Label3(0).Caption = "Work hours"

    Label1(13).Caption = "By"
    Label2(2).Caption = "Curr rec"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

    btnQuery.Caption = "Search"

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is Me.TxtMoveTo Then
        
        Else

            If Me.TxtModFlg.text = "R" Then
                btnNew_Click
            Else
                SendKeys "{TAB}"
            End If
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

Private Sub Form_Load()
    Dim My_SQL As String
    My_SQL = "  select Fullcode,Project_name from Projects"
    fill_combo DataCombo2, My_SQL

    Dim i As Integer

    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    On Error GoTo ErrTrap

    Resize_Form Me
    'load tblEmployee -----------------------------------------------
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees DCEmp_Name, True
    Set cSearch = New clsDCboSearch
    Set cSearch.Client = DCEmp_Name
    Dcombos.GetUsers DCUser
    Dcombos.Getsheft CboShift

    'DTTime.Value = Time
    With Me.CmbTime
        .AddItem "’"
        .ItemData(.NewIndex) = 0
        .AddItem "„"
        .ItemData(.NewIndex) = 1
    End With

    '----------------------------------------------------------------------------
    With CmbTimeType
        .AddItem "⁄„·"
        .ItemData(.NewIndex) = 0
        .AddItem "⁄ÿ·…"
        .ItemData(.NewIndex) = 1
    End With

    SetDtpickerDate DtDate
    SetDtpickerDate DtpPresent
    SetDtpickerDate Me.DtpDeparture

    With Me.DtpPresentHour
        .Format = dtpCustom
        .CustomFormat = "'Time: 'hh:mm tt"
        .value = Time
    End With

    With Me.DtpDepHour
        .Format = dtpCustom
        .CustomFormat = "'Time: 'hh:mm tt"
        .value = Time
    End With

    'With Me.CboShift2
    '    .Clear
    '    .AddItem "«·œÊ«„ «·√Ê·"
    '    .AddItem "«·œÊ«„ «·À«‰Ï"
    '    .AddItem "«·œÊ«„ «·À«·À"
    'End With
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        .WallPaper = GrdBack.Picture
        .Rows = .FixedRows + 3
        .TextMatrix(0, .ColIndex("ShiftNO")) = "«·œÊ«„"
        .TextMatrix(0, .ColIndex("HoursCount")) = "⁄œœ ”«⁄«  «·⁄„·"
        .TextMatrix(1, .ColIndex("ShiftNO")) = "«·œÊ«„ «·√Ê·"
        .TextMatrix(2, .ColIndex("ShiftNO")) = "«·œÊ«„ «·À«‰Ï"
        .TextMatrix(3, .ColIndex("ShiftNO")) = "«·œÊ«„ «·À«·À"
    End With

    My_SQL = "select * From tblPresentTime where Present_Type=0"
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdText

    DTDate_Click
    BtnFirst_Click
    Me.TxtModFlg.text = "R"
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

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

                'btnSave
            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Terminate()
    Set FrmVacancy = Nothing

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

    Set TTP = Nothing
ErrTrap:
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("tblPresentTime", "Present_ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("Present_ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("Emp_ID").value = IIf(DCEmp_Name.text <> "", Trim(DCEmp_Name.BoundText), Null)
    RsSavRec.Fields("UserID").value = Me.DCUser.BoundText
    RsSavRec.Fields("PresentDate").value = IIf(DtDate.value <> "", Trim(DtDate.value), Null)
    RsSavRec.Fields("Present_Type").value = 0
    RsSavRec.Fields("Present_Code").value = IIf(TXTCode.text <> "", Trim(TXTCode.text), Null)
    RsSavRec.Fields("IntervalNO").value = IIf(CboShift.text <> "", Trim(CboShift.BoundText), Null)

    RsSavRec("GenPresentTime").value = GetPresentTime
    RsSavRec("CurrentPresentTime").value = FormatDateTime(Me.lbl(8).Caption, vbLongTime)
    RsSavRec("CurrentDepartureTime").value = FormatDateTime(Me.lbl(10).Caption, vbLongTime)
    RsSavRec("LateTime").value = Me.lbl(12).Caption
    RsSavRec("CurrentWorkMints").value = 600

    If Me.lbl(12).ForeColor = vbRed Then
        RsSavRec("LateTimeDiscountValue").value = CalculateLateDiscount(Me.DCEmp_Name.BoundText, Trim(Me.lbl(12).Caption))
    Else
        RsSavRec("LateTimeDiscountValue").value = 0
    End If

    If Me.Chk.value = vbChecked Then
        RsSavRec("GenDepartureTime").value = GetDepTime
        RsSavRec("WorkHoursCount").value = ConvertHoursToMints(Me.lbl(4).Caption)
    Else
        RsSavRec("GenDepartureTime").value = Null
        RsSavRec("WorkHoursCount").value = Null
    End If

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ .", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()
    Dim i As Integer
    Dim M_Date As Date
    Dim m_Temp As Date

    On Error GoTo ErrTrap

    TxtPresentTime_ID.text = IIf(IsNull(RsSavRec.Fields("Present_ID").value), "", RsSavRec.Fields("Present_ID").value)
    DCEmp_Name.BoundText = IIf(IsNull(RsSavRec.Fields("Emp_ID").value), "", RsSavRec.Fields("Emp_ID").value)
    DtDate.value = IIf(IsNull(RsSavRec.Fields("PresentDate").value), "", RsSavRec.Fields("PresentDate").value)
    Me.DCUser.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    TXTCode.text = IIf(IsNull(RsSavRec.Fields("Present_Code").value), "", RsSavRec.Fields("Present_Code").value)
    TxtEmp_Code.text = GetEmpCode("Emp_Code", " Emp_ID=" & val(DCEmp_Name.BoundText))
    CboShift.BoundText = IIf(IsNull(RsSavRec.Fields("IntervalNO").value), "", RsSavRec.Fields("IntervalNO").value)

    If Not IsNull(RsSavRec("CurrentPresentTime").value) Then
        Me.lbl(13).Caption = RsSavRec("CurrentPresentTime").value
    Else
        Me.lbl(13).Caption = ""
    End If

    If Not IsNull(RsSavRec("CurrentDepartureTime").value) Then
        Me.lbl(15).Caption = RsSavRec("CurrentDepartureTime").value
    Else
        Me.lbl(15).Caption = ""
    End If

    If Not IsNull(RsSavRec("GenPresentTime").value) Then
        M_Date = RsSavRec("GenPresentTime").value
        Me.DtpPresent.value = FormatDateTime(M_Date, vbShortDate)
        m_Temp = FormatDateTime(M_Date, vbShortTime)
        Me.DtpPresentHour.value = m_Temp
    End If

    If Not IsNull(RsSavRec("GenDepartureTime").value) Then
        Me.Chk.value = vbChecked
        M_Date = RsSavRec("GenDepartureTime").value
        Me.DtpDeparture.value = FormatDateTime(M_Date, vbShortDate)
        m_Temp = FormatDateTime(M_Date, vbShortTime)
        Me.DtpDepHour.value = m_Temp
    Else
        Me.Chk.value = vbUnchecked
    End If

    Chk_Click
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    DTDate_Click
ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecID As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Tim_Timer()

    If SystemOptions.UserInterface = ArabicInterface Then
        Me.LblZoneSetting.Caption = "«· «—ÌŒ «·Õ«·Ï ›Ï «·ÃÂ«“ : " & FormatDateTime(Date, vbLongDate)
    Else
        Me.LblZoneSetting.Caption = "Current System Date " & FormatDateTime(Date, vbLongDate)
    End If

End Sub

Private Sub TxtEmp_Code_KeyUp(KeyCode As Integer, _
                              Shift As Integer)
    DCEmp_Name.BoundText = GetEmpCode("Emp_ID", " Emp_Code='" & TxtEmp_Code.text & "'")

End Sub

Private Sub TxtMoveTo_KeyPress(KeyAscii As Integer)
    Dim m_BookMark As Variant
    Dim Msg As String

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0

        If val(Me.TxtMoveTo.text) <> 0 Then
            If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
                m_BookMark = RsSavRec.Bookmark
                RsSavRec.find "Present_ID=" & val(Me.TxtMoveTo.text), , adSearchForward, ADODB.BookmarkEnum.adBookmarkFirst

                If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
                    FiLLTXT
                Else
                    Msg = "·„ Ì „ «·⁄ÀÊ— ⁄·Ï Â–« «·”Ã·...!!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    RsSavRec.Bookmark = m_BookMark
                End If
            End If
        End If
    End If

End Sub

Private Sub TxtPresentTime_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = TxtMod
End Sub

Public Function FindRec(ByVal RecID As Long)
    On Error GoTo ErrTrap

    RsSavRec.find "Present_ID=" & RecID, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
    
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
    
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    
        ' btnNext.Enabled = False
        ' btnPrevious.Enabled = False
        ' btnFirst.Enabled = False
        ' btnLast.Enabled = False
        FrmBrngTime.Enabled = True
    ElseIf TxtModFlg.text = "R" Then
        FrmBrngTime.Enabled = False
        btnModify.Enabled = True
        btnDelete.Enabled = True

        If TxtPresentTime_ID.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If
    
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.text = "E" Then
        FrmBrngTime.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
    
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
    
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Private Sub GetTimeDetails()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    If Weekday(DtDate.value) = 7 Then
        My_SQL = "select * From tblTimeSetting where DayNO=" & 1 & ""
    Else
        My_SQL = "select * From tblTimeSetting where DayNO=" & Weekday(DtDate.value) + 1 & ""
    End If

    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
    
        CmbTimeType.ListIndex = IIf(IsNull(rs.Fields("Is_WorkDay").value), -1, rs.Fields("Is_WorkDay").value)
       
        TxtHour.text = IIf(IsNull(rs.Fields("Bring_HourTime").value), "", rs.Fields("Bring_HourTime").value)
                                    
        TxtMinute.text = IIf(IsNull(rs.Fields("Bring_MinuteTime").value), "", rs.Fields("Bring_MinuteTime").value)
    
        CmbTime.ListIndex = IIf(IsNull(rs.Fields("Bring_Time").value), -1, rs.Fields("Bring_Time").value)
    
    End If

    'If CmbTimeType.ListIndex = 0 Then
    '    FrmBrngTime.Enabled = True
    'Else
    '    FrmBrngTime.Enabled = False
    '    TxtEmp_Code.Text = ""
    '    DCEmp_Name.Text = ""
    'End If
    LabWork.Caption = CmbTimeType.text

    rs.Close
    Set rs = Nothing
ErrTrap:
End Sub

Private Function GetEmpCode(ByVal Fild As String, _
                            ByVal whr As String) As String
    On Error GoTo ErrTrap
    Dim My_SQL As String

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    GetEmpCode = ""
    My_SQL = "select " & Fild & " From TblEmployee where " & whr
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        GetEmpCode = IIf(IsNull(rs.Fields(Fild).value), "", rs.Fields(Fild).value)
    End If

    rs.Close
    Set rs = Nothing
ErrTrap:
End Function

Private Function ChkEmpExist() As Boolean
    On Error GoTo ErrTrap
    Dim My_SQL As String

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    ChkEmpExist = False
    My_SQL = "select * From tblPresentTime where Present_type=0 and Emp_ID=" & DCEmp_Name.BoundText & " and  CONVERT (nvarchar(50),GenPresentTime ,101)=" & SQLDate(DtpPresent.value, True) & ""

    If Me.TxtModFlg.text = "E" Then
        My_SQL = My_SQL + " and Present_ID  <> " & val(TxtPresentTime_ID.text) & """"
    End If
            
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        ChkEmpExist = True
    End If

    rs.Close
    Set rs = Nothing
ErrTrap:
End Function

Private Function ChkEmpComeToday() As Boolean
    On Error GoTo ErrTrap
    Dim My_SQL As String

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    ChkEmpComeToday = True
    My_SQL = "select * From QryAbsentEmp where AbsDate=" & SQLDate(DtDate.value, True) & " and Emp_ID=" & DCEmp_Name.BoundText
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        ChkEmpComeToday = False
    End If

    rs.Close
    Set rs = Nothing
ErrTrap:
End Function

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap

    Dim Wrap As String
    Dim Msg As String
    Dim BolRtl As Boolean
    Wrap = Chr(13) + Chr(10)

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = Fra(4).Caption & Wrap & "Â‰«  ﬁÊ„ » ÕœÌœ «·ÌÊ„ Ê ”«⁄… Õ÷Ê— «·„ÊŸ›"
        .AddControl Fra(4), Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = Fra(4).Caption & Wrap & " ÕœÌœ ÌÊ„ «·Õ÷Ê—"
        .AddControl DtpPresent, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = Fra(4).Caption & Wrap & " ÕœÌœ ”«⁄… «·Õ÷Ê—"
        .AddControl DtpPresentHour, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = Chk.Caption & Wrap & "≈–« ﬂ‰   —Ìœ  ”ÃÌ· ≈‰’—«› «·„ÊŸ› „‰ Â–Â «·‘«‘… " & Wrap & "›«‰Â Ì„ﬂ‰ﬂ Â–« „‰ Œ·«·  ›⁄Ì· Â–« «·ŒÌ«—.."
        .AddControl Chk, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = Chk.Caption & Wrap & " ÕœÌœ ÌÊ„ «·√‰’—«›"
        .AddControl DtpDeparture, Msg, BolRtl
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630, BolRtl
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = Chk.Caption & Wrap & " ÕœÌœ ”«⁄… «·√‰’—«›"
        .AddControl DtpDepHour, Msg, BolRtl
    End With

ErrTrap:

End Sub

Private Function GetPresentTime() As Date
    Dim GetPreDate As Date
    Dim StrTemp As String
    StrTemp = Me.DtpPresent.value & " " & FormatDateTime(Me.DtpPresentHour.value, vbShortTime)
    GetPreDate = FormatDateTime(StrTemp, vbGeneralDate)
    GetPresentTime = GetPreDate
End Function

Private Function GetDepTime()
    Dim GenDepDate As Date
    Dim GetPreDate As Date
    Dim StrTemp As String

    If Me.Chk.value = vbChecked Then
        StrTemp = Me.DtpDeparture.value & " " & FormatDateTime(Me.DtpDepHour.value, vbShortTime)
        GenDepDate = FormatDateTime(StrTemp, vbGeneralDate)
        GetDepTime = GenDepDate
    End If

End Function

Private Sub GetEmpSetting()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim m_Temp As Date
    Dim StrTemp As String

    If Me.DCEmp_Name.BoundText = "" Then
        TxtEmp_Code.text = ""
    Else
        '«·√⁄œ«œ«  «·Œ«’… »«·„ÊŸ›
        TxtEmp_Code.text = GetEmpCode("Emp_Code", " Emp_ID=" & val(DCEmp_Name.BoundText))
        StrSQL = "Select * From tblTimeSettingEmp "
        StrSQL = StrSQL + " Where Emp_ID=" & val(DCEmp_Name.BoundText) & ""
        StrSQL = StrSQL + " AND DayNO=" & Weekday(DtpPresent.value, vbSaturday) & ""
        StrSQL = StrSQL + " AND Is_WorkDay=0"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If (rs.BOF Or rs.EOF) Then
            '«·√⁄œ«œ«  «·Œ«’… »«·‘—ﬂ…
            StrSQL = "Select * From tblTimeSetting "
            StrSQL = StrSQL + " Where DayNO=" & Weekday(DtpPresent.value, vbSaturday) & ""
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        End If

        If Not (rs.BOF Or rs.EOF) Then
            If Not IsNull(rs("Bring_HourTime").value) Then
                If rs("Bring_Time").value = 0 Then
                    StrTemp = rs("Bring_HourTime").value & ":" & rs("Bring_MinuteTime").value & " AM"
                Else
                    StrTemp = rs("Bring_HourTime").value & ":" & rs("Bring_MinuteTime").value & " PM"
                End If

                m_Temp = FormatDateTime(StrTemp, vbLongTime)
                Me.lbl(8).Caption = m_Temp
            End If

            If Not IsNull(rs("Go_Time").value) Then
                If rs("Go_Time").value = 0 Then
                    StrTemp = rs("Go_HourTime").value & ":" & rs("Go_MinuteTime").value & " AM"
                Else
                    StrTemp = rs("Go_HourTime").value & ":" & rs("Go_MinuteTime").value & " PM"
                End If

                m_Temp = FormatDateTime(StrTemp, vbLongTime)
                Me.lbl(10).Caption = m_Temp
            End If
        End If
    End If

End Sub

Private Sub CalculateTimes()
    Dim m_PreDate As Date
    Dim IntMintsCount As Double
    Dim IntHoursCount As Double
    Dim IntMin As Integer
    Dim StrSing As String

    Dim m_DateCome As Date
    Dim m_DateOut As Date

    Dim StrTemp As String

    If Me.lbl(8).Caption <> "" Then
        If IsDate(Me.lbl(8).Caption) Then
            m_PreDate = CDate(Me.lbl(8).Caption)
            IntMintsCount = DateDiff("n", m_PreDate, Me.DtpPresentHour.value)
            IntHoursCount = IntMintsCount \ 60
            IntMin = IntMintsCount Mod 60

            If IntHoursCount < 0 Or IntMin < 0 Then
                'Stop
                StrSing = "-"
                IntHoursCount = Abs(IntHoursCount)
                IntMin = Abs(IntMin)
                Me.lbl(12).ForeColor = &H8000&
            Else
                StrSing = ""
                Me.lbl(12).ForeColor = vbRed
            End If

            StrTemp = StrSing & Format(IntHoursCount, "00") & ":" & Format(IntMin, "00")
            lbl(12).Caption = StrTemp
        End If
    End If

    If Me.Chk.value = vbChecked Then
        m_DateCome = GetPresentTime
        m_DateOut = GetDepTime
        IntMintsCount = DateDiff("n", m_DateCome, m_DateOut)
        IntHoursCount = IntMintsCount \ 60
        IntMin = IntMintsCount Mod 60

        If IntHoursCount < 0 Or IntMin < 0 Then
            'Stop
            StrTemp = "Œÿ√"
            Me.lbl(4).ForeColor = vbRed
        Else
            StrSing = ""
            StrTemp = StrSing & Format(IntHoursCount, "00") & ":" & Format(IntMin, "00")
            Me.lbl(4).ForeColor = &H8000&
        End If

        lbl(4).Caption = StrTemp
    End If

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)

    If Lngid <> 0 Then
        If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
            RsSavRec.find "Present_ID=" & Lngid, , adSearchForward, 1

            If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
                FiLLTXT
            End If
        End If
    End If

End Sub

