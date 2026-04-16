VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmDeported 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmDeported.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   14550
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
      TabIndex        =   37
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmDeported.frx":6852
      Left            =   15480
      List            =   "FrmDeported.frx":6862
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   36
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
      TabIndex        =   30
      Top             =   0
      Width           =   14625
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   450
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
         ButtonImage     =   "FrmDeported.frx":687B
         ColorButton     =   16777215
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   915
         TabIndex        =   32
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
         ButtonImage     =   "FrmDeported.frx":6C15
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1515
         TabIndex        =   33
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
         ButtonImage     =   "FrmDeported.frx":6FAF
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   2040
         TabIndex        =   34
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
         ButtonImage     =   "FrmDeported.frx":7349
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÃœÊ·  —ÕÌ·"
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
         TabIndex        =   35
         Top             =   240
         Width           =   4080
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   13200
         Picture         =   "FrmDeported.frx":76E3
         Stretch         =   -1  'True
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frm2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   6975
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   720
      Width           =   14595
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   3375
         Left            =   120
         TabIndex        =   69
         Top             =   3600
         Width           =   14415
         Begin VB.TextBox TxtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   915
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   93
            Top             =   2400
            Width           =   12975
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  «·„‘—›"
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   1200
            Width           =   14295
            Begin VB.TextBox TxtSuperVisName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   240
               Width           =   7095
            End
            Begin VB.TextBox TxtPhone 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   240
               Width           =   1545
            End
            Begin VB.TextBox TxtAddress 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5280
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   87
               Top             =   600
               Width           =   7095
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   11310
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   600
               Visible         =   0   'False
               Width           =   1065
            End
            Begin MSDataListLib.DataCombo DcbSuperID 
               Bindings        =   "FrmDeported.frx":8AE8
               Height          =   315
               Left            =   5280
               TabIndex        =   88
               Top             =   600
               Visible         =   0   'False
               Width           =   6015
               _ExtentX        =   10610
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
               Caption         =   "«·Â« ›"
               Height          =   285
               Index           =   16
               Left            =   3450
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   240
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄‰Ê«‰ «·ÊÃÂ…"
               Height          =   285
               Index           =   15
               Left            =   12840
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   600
               Width           =   1275
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„‘—›"
               Height          =   285
               Index           =   14
               Left            =   12840
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   240
               Width           =   1275
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„—Õ·"
            Height          =   1095
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   3120
            Visible         =   0   'False
            Width           =   8535
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6030
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox TxtRelayName2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   600
               Width           =   6855
            End
            Begin XtremeSuiteControls.RadioButton ChRelayType2 
               Height          =   255
               Index           =   0
               Left            =   6840
               TabIndex        =   76
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
            Begin MSDataListLib.DataCombo DcbRelayID2 
               Bindings        =   "FrmDeported.frx":8AFD
               Height          =   315
               Left            =   240
               TabIndex        =   77
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
            Begin XtremeSuiteControls.RadioButton ChRelayType2 
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
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E2E9E9&
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   120
            Width           =   14295
            Begin VB.TextBox TxtDiffTime 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   240
               Visible         =   0   'False
               Width           =   1545
            End
            Begin VB.TextBox TxtDiffDate 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   240
               Visible         =   0   'False
               Width           =   1545
            End
            Begin MSDataListLib.DataCombo DcbLocatioID2 
               Bindings        =   "FrmDeported.frx":8B12
               Height          =   315
               Left            =   5520
               TabIndex        =   71
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
            Begin MSComCtl2.DTPicker CurrDate2 
               Height          =   315
               Left            =   10920
               TabIndex        =   80
               Top             =   240
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   556
               _Version        =   393216
               Format          =   97583105
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal CurrDateH2 
               Height          =   315
               Left            =   9480
               TabIndex        =   81
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
            End
            Begin MSComCtl2.DTPicker TimeIn 
               Height          =   315
               Left            =   2280
               TabIndex        =   83
               Top             =   600
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Format          =   97583106
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   10
               Left            =   7440
               TabIndex        =   95
               Top             =   240
               Width           =   1965
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "“„‰ «·Ê’Ê·"
               Height          =   285
               Index           =   6
               Left            =   4170
               TabIndex        =   84
               Top             =   615
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ"
               Height          =   285
               Index           =   17
               Left            =   13170
               TabIndex        =   82
               Top             =   255
               Width           =   885
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Êﬁ⁄"
               Height          =   285
               Index           =   12
               Left            =   13050
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   600
               Width           =   1005
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   285
            Index           =   9
            Left            =   13560
            TabIndex        =   94
            Top             =   2760
            Width           =   885
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   2535
         Left            =   0
         TabIndex        =   29
         Top             =   720
         Width           =   14535
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Height          =   2535
            Left            =   120
            TabIndex        =   57
            Top             =   120
            Width           =   14415
            Begin VB.Frame Frame11 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   1680
               Width           =   14295
               Begin MSDataListLib.DataCombo DcbFromCity 
                  Bindings        =   "FrmDeported.frx":8B27
                  Height          =   315
                  Left            =   9600
                  TabIndex        =   13
                  Top             =   720
                  Visible         =   0   'False
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
               Begin MSDataListLib.DataCombo DcbToCity 
                  Bindings        =   "FrmDeported.frx":8B3C
                  Height          =   315
                  Left            =   5640
                  TabIndex        =   14
                  Top             =   720
                  Visible         =   0   'False
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
               Begin MSComCtl2.DTPicker TimeOut 
                  Height          =   315
                  Left            =   2040
                  TabIndex        =   12
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   97583106
                  CurrentDate     =   38784
               End
               Begin MSDataListLib.DataCombo DcbPath 
                  Bindings        =   "FrmDeported.frx":8B51
                  Height          =   315
                  Left            =   5640
                  TabIndex        =   97
                  Top             =   240
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "“„‰ «·„€«œ—…"
                  Height          =   285
                  Index           =   27
                  Left            =   3810
                  TabIndex        =   65
                  Top             =   255
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ï"
                  Height          =   285
                  Index           =   25
                  Left            =   8640
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   285
                  Index           =   24
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   435
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„”«—« "
                  Height          =   285
                  Index           =   23
                  Left            =   12960
                  RightToLeft     =   -1  'True
                  TabIndex        =   62
                  Top             =   240
                  Width           =   795
               End
            End
            Begin VB.Frame Frame8 
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„—Õ·"
               Height          =   1095
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   3000
               Visible         =   0   'False
               Width           =   8535
               Begin VB.TextBox TxtRelayName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   600
                  Width           =   6855
               End
               Begin VB.TextBox TxtSwapCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6030
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   240
                  Width           =   1065
               End
               Begin XtremeSuiteControls.RadioButton ChRelayType 
                  Height          =   255
                  Index           =   0
                  Left            =   6840
                  TabIndex        =   7
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
               Begin MSDataListLib.DataCombo RelayID 
                  Bindings        =   "FrmDeported.frx":8B66
                  Height          =   315
                  Left            =   240
                  TabIndex        =   9
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
               Begin XtremeSuiteControls.RadioButton ChRelayType 
                  Height          =   255
                  Index           =   1
                  Left            =   6840
                  TabIndex        =   10
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
               Height          =   1095
               Left            =   5400
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   600
               Width           =   8775
               Begin VB.TextBox TxtSearchCode 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   6030
                  RightToLeft     =   -1  'True
                  TabIndex        =   3
                  Top             =   240
                  Width           =   1065
               End
               Begin VB.TextBox TxtDriverName 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   240
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   600
                  Width           =   6855
               End
               Begin XtremeSuiteControls.RadioButton ChTypeDrive 
                  Height          =   255
                  Index           =   0
                  Left            =   6840
                  TabIndex        =   2
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
                  Bindings        =   "FrmDeported.frx":8B7B
                  Height          =   315
                  Left            =   240
                  TabIndex        =   4
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
                  TabIndex        =   5
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
            Begin MSComCtl2.DTPicker CurrDate 
               Height          =   315
               Left            =   1800
               TabIndex        =   66
               Top             =   240
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               Format          =   97583105
               CurrentDate     =   38784
            End
            Begin Dynamic_Byte.NourHijriCal CurrDateH 
               Height          =   315
               Left            =   240
               TabIndex        =   67
               Top             =   240
               Width           =   1455
               _ExtentX        =   2778
               _ExtentY        =   556
            End
            Begin MSDataListLib.DataCombo DcbLocatioID 
               Bindings        =   "FrmDeported.frx":8B90
               Height          =   315
               Left            =   4200
               TabIndex        =   100
               Top             =   240
               Width           =   3615
               _ExtentX        =   6376
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
            Begin MSDataListLib.DataCombo DcbEqupID 
               Bindings        =   "FrmDeported.frx":8BA5
               Height          =   315
               Left            =   8880
               TabIndex        =   102
               Top             =   240
               Width           =   4215
               _ExtentX        =   7435
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
            Begin MSDataListLib.DataCombo DcbProgramID 
               Bindings        =   "FrmDeported.frx":8BBA
               Height          =   315
               Left            =   240
               TabIndex        =   104
               Top             =   1440
               Visible         =   0   'False
               Width           =   3615
               _ExtentX        =   6376
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
            Begin MSDataListLib.DataCombo DcbTypeTrip 
               Bindings        =   "FrmDeported.frx":8BCF
               Height          =   315
               Left            =   240
               TabIndex        =   106
               Top             =   720
               Width           =   3615
               _ExtentX        =   6376
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
            Begin MSDataListLib.DataCombo DcbTypePath 
               Bindings        =   "FrmDeported.frx":8BE4
               Height          =   315
               Left            =   240
               TabIndex        =   113
               Top             =   1200
               Width           =   3615
               _ExtentX        =   6376
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
               Caption         =   "‰Ê⁄ «·—Õ·…"
               Height          =   285
               Index           =   1
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   720
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·„”«—"
               Height          =   285
               Index           =   13
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   1200
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ«›·…"
               Height          =   285
               Index           =   3
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Êﬁ⁄"
               Height          =   285
               Index           =   5
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   240
               Width           =   1515
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Height          =   285
               Index           =   11
               Left            =   240
               TabIndex        =   96
               Top             =   240
               Width           =   765
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ"
               Height          =   285
               Index           =   0
               Left            =   3000
               TabIndex        =   68
               Top             =   255
               Width           =   1515
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   14535
         Begin VB.TextBox TxtPrintCount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   122
            Top             =   0
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12840
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtNoteSerialOrder 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   118
            Top             =   240
            Width           =   1335
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmDeported.frx":8BF9
            Left            =   2520
            List            =   "FrmDeported.frx":8BFB
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   240
            Width           =   1470
         End
         Begin VB.TextBox TxtHajzNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   111
            Top             =   720
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox TxtOrderID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   360
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   98
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   10920
            TabIndex        =   0
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   97583105
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "FrmDeported.frx":8BFD
            Height          =   315
            Left            =   7080
            TabIndex        =   1
            Top             =   240
            Width           =   2055
            _ExtentX        =   3625
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
            Left            =   9720
            TabIndex        =   58
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
         End
         Begin MSDataListLib.DataCombo SeasonsID 
            Height          =   315
            Left            =   4680
            TabIndex        =   119
            Top             =   240
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   330
            Index           =   30
            Left            =   6255
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   240
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   285
            Index           =   21
            Left            =   3840
            TabIndex        =   117
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÕÃ“"
            Height          =   285
            Index           =   20
            Left            =   480
            TabIndex        =   112
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„  «ﬂÌœ ÕÃ“"
            Height          =   285
            Index           =   19
            Left            =   1440
            TabIndex        =   99
            Top             =   240
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   285
            Index           =   7
            Left            =   8880
            TabIndex        =   56
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ "
            Height          =   285
            Index           =   4
            Left            =   13680
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            Height          =   285
            Index           =   2
            Left            =   12210
            TabIndex        =   27
            Top             =   255
            Width           =   645
         End
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì«‰«  «·Ê’Ê·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   18
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   3240
         Width           =   4515
      End
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      RightToLeft     =   -1  'True
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   38
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
      TabIndex        =   39
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
      Height          =   1305
      Left            =   0
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   7680
      Width           =   14595
      _cx             =   25744
      _cy             =   2302
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
         TabIndex        =   42
         Top             =   0
         Width           =   3855
         Begin VB.Label LabCountRec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
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
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   735
         Left            =   0
         TabIndex        =   41
         Top             =   480
         Width           =   14175
         Begin ImpulseButton.ISButton btnNew 
            Height          =   330
            Left            =   12720
            TabIndex        =   17
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
            ButtonImage     =   "FrmDeported.frx":8C12
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnSave 
            Height          =   330
            Left            =   9360
            TabIndex        =   19
            ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            ButtonImage     =   "FrmDeported.frx":F474
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnModify 
            Height          =   330
            Left            =   11160
            TabIndex        =   18
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
            ButtonImage     =   "FrmDeported.frx":F80E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUndo 
            Height          =   330
            Left            =   7800
            TabIndex        =   20
            ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
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
            ButtonImage     =   "FrmDeported.frx":16070
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnDelete 
            Height          =   330
            Left            =   6600
            TabIndex        =   21
            ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
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
            ButtonImage     =   "FrmDeported.frx":1640A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   330
            Left            =   120
            TabIndex        =   22
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
            ButtonImage     =   "FrmDeported.frx":169A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   405
            Left            =   2640
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â «·ﬂ·"
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
            ButtonImage     =   "FrmDeported.frx":16D3E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   330
            Left            =   1680
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   240
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmDeported.frx":1D5A0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   5160
            TabIndex        =   109
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â Œ—ÊÃ"
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
            ButtonImage     =   "FrmDeported.frx":1D93A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   405
            Left            =   3840
            TabIndex        =   110
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   240
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â Ê’Ê·"
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
            ButtonImage     =   "FrmDeported.frx":2419C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   315
         Left            =   8400
         TabIndex        =   47
         Top             =   120
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   570
         Left            =   120
         TabIndex        =   51
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
         Caption         =   "«·„” Œœ„"
         Height          =   270
         Index           =   8
         Left            =   13080
         TabIndex        =   48
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
            Picture         =   "FrmDeported.frx":2A9FE
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDeported.frx":2AD98
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDeported.frx":2B132
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDeported.frx":2B4CC
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDeported.frx":2B866
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDeported.frx":2BC00
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDeported.frx":2BF9A
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDeported.frx":2C534
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   49
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
      ButtonImage     =   "FrmDeported.frx":2C8CE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   52
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
      ButtonImage     =   "FrmDeported.frx":33130
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   53
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
      ButtonImage     =   "FrmDeported.frx":39992
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
      TabIndex        =   50
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmDeported"
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
 Dim II As Long

Private Sub CBoBasedON_Change()
If val(CBoBasedON.ListIndex) = 0 Then
TxtNoteSerialOrder.Visible = False
TxtOrderID.Visible = False
TxtHajzNo.Visible = False
lbl(19).Visible = False
lbl(20).Visible = False
loadcombo2
Else
TxtNoteSerialOrder.Visible = True
TxtOrderID.Visible = True
TxtHajzNo.Visible = True
lbl(19).Visible = True
lbl(20).Visible = True
End If
End Sub

Private Sub CBoBasedON_Click()
CBoBasedON_Change
End Sub

Private Sub ChRelayType_Click(Index As Integer)
TxtSwapCode.Enabled = False
RelayID.Enabled = False
TxtRelayName.Enabled = False
If ChRelayType(0).value = True Then
TxtSwapCode.Enabled = True
RelayID.Enabled = True
TxtRelayName.Text = ""
Else
TxtRelayName.Enabled = True
TxtSwapCode.Text = ""
RelayID.BoundText = 0
End If
End Sub

Private Sub ChRelayType2_Click(Index As Integer)
Text2.Enabled = False
DcbRelayID2.Enabled = False
TxtRelayName2.Enabled = False
If ChRelayType2(0).value = True Then
Text2.Enabled = True
DcbRelayID2.Enabled = True
TxtRelayName2.Text = ""
Else
TxtRelayName2.Enabled = True
Text2.Text = ""
DcbRelayID2.BoundText = 0
End If
End Sub

Private Sub ChTypeDrive_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
TxtSearchCode.Enabled = False
DcbDriverID.Enabled = False
TxtDriverName.Enabled = False
If ChTypeDrive(0).value = True Then
DcbDriverID.Enabled = True
TxtSearchCode.Enabled = True
TxtDriverName.Text = ""
Else
TxtDriverName.Enabled = True
DcbDriverID.BoundText = 0
TxtSearchCode.Text = ""
End If
End If
End Sub



Private Sub CurrDate_Change()
If Me.TxtModFlg.Text <> "R" Then
         CurrDateH.value = ToHijriDate(CurrDate.value)
End If
If SystemOptions.UserInterface = ArabicInterface Then
Select Case Weekday(CurrDate.value)
Case 1
         lbl(11).Caption = "«·«Õœ "
Case 2
         lbl(11).Caption = "«·«À‰Ì‰ "
Case 3
         lbl(11).Caption = "«·À·«À«¡ "
Case 4
         lbl(11).Caption = "«·«—»⁄«¡ "
Case 5
         lbl(11).Caption = "«·Œ„Ì” "
Case 6
         lbl(11).Caption = "«·Ã„⁄… "
Case 7
         lbl(11).Caption = "«·”»  "
 End Select
Else
lbl(11).Caption = WeekdayName(Weekday(CurrDate.value))
End If
End Sub

Private Sub CurrDate2_Change()
If Me.TxtModFlg.Text <> "R" Then
         CurrDateH2.value = ToHijriDate(CurrDate2.value)
End If
If SystemOptions.UserInterface = ArabicInterface Then
Select Case Weekday(CurrDate2.value)
Case 1
         lbl(10).Caption = "«·«Õœ "
Case 2
         lbl(10).Caption = "«·«À‰Ì‰ "
Case 3
         lbl(10).Caption = "«·À·«À«¡ "
Case 4
         lbl(10).Caption = "«·«—»⁄«¡ "
Case 5
         lbl(10).Caption = "«·Œ„Ì” "
Case 6
         lbl(10).Caption = "«·Ã„⁄… "
Case 7
         lbl(10).Caption = "«·”»  "
 End Select
Else
lbl(10).Caption = WeekdayName(Weekday(CurrDate2.value))
End If
ClCulteDiffDateTime
End Sub



Private Sub CurrDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    CurrDate.value = ToGregorianDate(CurrDateH.value)
    End If
End Sub
Private Sub CurrDateH2_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    CurrDate2.value = ToGregorianDate(CurrDateH2.value)
    End If
End Sub
Private Sub DcbDriverID_Change()
Dim CarID As Double
If val(DcbDriverID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbDriverID.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
   If Me.TxtModFlg.Text <> "R" Then
 If val(DcbDriverID.BoundText) <> 0 Then
   GetCarInfo val(DcbDriverID.BoundText), CarID, 0
   DcbEqupID.BoundText = CarID
   End If
   Exit Sub
 End If
End Sub
Private Sub DcbDriverID_Click(Area As Integer)
DcbDriverID_Change
End Sub

Private Sub DcbDriverID_KeyPress(KeyAscii As Integer)
Dim CarID As Double
If KeyAscii = vbKeyReturn Then
   If Me.TxtModFlg.Text <> "R" Then
   GetCarInfo val(DcbDriverID.BoundText), CarID, 0
   DcbEqupID.BoundText = CarID
   End If
 End If
End Sub

Private Sub DcbEqupID_Change()
DcbEqupID_Click (0)
End Sub

Private Sub DcbEqupID_Click(Area As Integer)
Dim EmpID As Double
 If Me.TxtModFlg.Text <> "R" Then
 If val(DcbEqupID.BoundText) <> 0 Then
   GetCarInfo EmpID, val(DcbEqupID.BoundText), 1
   Me.DcbDriverID.BoundText = EmpID
   Exit Sub
 End If
   End If
End Sub
Sub GetCarInfo(Optional ByRef Emp_id As Double, Optional ByRef CarID As Double, Optional Type1 As Integer = 0)
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
Dim Sql As String
Sql = "Select      id, Emp_id"
Sql = Sql & " FROM         dbo.TblCarsData"
If Type1 = 0 Then
Sql = Sql & " where Emp_id= " & Emp_id & ""
Else
Sql = Sql & " where ID= " & CarID & ""
End If
Rs5.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs5.RecordCount > 0 Then
Emp_id = IIf(IsNull(Rs5("Emp_id").value), 0, Rs5("Emp_id").value)
CarID = IIf(IsNull(Rs5("ID").value), 0, Rs5("ID").value)
Else
Emp_id = 0
CarID = 0
End If

End Sub

Private Sub DcbLocatioID2_Change()
DcbLocatioID2_Click (0)
End Sub

Private Sub DcbLocatioID2_Click(Area As Integer)
ClCulteDiffDateTime
End Sub

Private Sub DcbRelayID2_Change()
DcbRelayID2_Click (0)
End Sub

Private Sub DcbRelayID2_Click(Area As Integer)
If val(DcbRelayID2.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbRelayID2.BoundText, EmpCode
    Text2.Text = EmpCode
End Sub

Private Sub DcbSuperID_Change()
DcbSuperID_Click (0)
End Sub

Private Sub DcbSuperID_Click(Area As Integer)
If val(DcbSuperID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcbSuperID.BoundText, EmpCode
    Text3.Text = EmpCode
If Me.TxtModFlg.Text <> "R" Then
GetEmp val(DcbSuperID.BoundText)
End If
End Sub

Sub ladData()
 Dim Dcombos As New ClsDataCombos
     Dcombos.ClearMyDataCombo DcbDriverID
     Dim str As String
If SystemOptions.UserInterface = ArabicInterface Then
    str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee ,dbo.TblEmployee.BranchId"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name ,dbo.TblEmployee.BranchId "
   End If
    ' If Me.TxtModFlg.Text <> "R" Then
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    str = str & "     where  (( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1)) "
    str = str & "  and  (dbo.TblEmployee.BranchId=0 or dbo.TblEmployee.BranchId is null or         dbo.TblEmployee.BranchId in(" & Current_branchSql & "))"
    'Else
  '      str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
   ' str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
   ' str = str & "     where  (( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1)) "
    'End If
   
    fill_combo DcbDriverID, str
End Sub

Private Sub ISButton2_Click()
print_report , 1

End Sub

Private Sub ISButton3_Click()
print_report , 2

End Sub

Private Sub ISButton8_Click()
  Unload FrmSearch_Hajj
         FrmSearch_Hajj.SendForm = "Deported"
         FrmSearch_Hajj.show
End Sub

Sub ClCulteDiffDateTime()
If Me.TxtModFlg.Text <> "R" Then
Dim X As Double
If val(DcbLocatioID2.BoundText) <> 0 And DcbLocatioID2.Text <> "" Then
TxtDiffDate.Text = DateDiff("d", CurrDate.value, CurrDate2.value)
X = Abs(DateDiff("n", TimeIn.value, TimeOut.value) / 60)
X = Fix(X)
  TxtDiffTime.Text = X & ":" & Abs(DateDiff("n", TimeIn.value, TimeOut.value) Mod 60)
End If
End If
End Sub

Private Sub RelayID_Change()
If val(RelayID.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , RelayID.BoundText, EmpCode
    TxtSwapCode.Text = EmpCode
End Sub
    Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String

     Dim str  As String
      str = "  select   e.Emp_ID Emp_ID , e.Emp_Name   Emp_Name  from TblEmployee e, TblEmpJobsTypes  j"
      str = str & "   Where e.JobTypeID = j.JobTypeID"
      str = str & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
    fill_combo DcbDriverID, str
         str = "  select   id, OperatorN from TblCarsData WHERE     ( NOT (OperatorN IS NULL)  AND OperatorN <>  '')  "
         str = str & "  and  (TblCarsData.Branch_NO=0 or TblCarsData.Branch_NO is null or    TblCarsData.Branch_NO in(" & Current_branchSql & "))"
    fill_combo DcbEqupID, str
   If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
  End If
  str = str & " where Omra_Hajj=0"
   fill_combo SeasonsID, str
       ladData
    conection = "select * from TblDeported order by  ID "
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open conection, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
    Dcombos.getCountriesGovernments Me.DcbFromCity
    Dcombos.getCountriesGovernments Me.DcbToCity
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetEmployees Me.RelayID
    Dcombos.GetEmployees Me.DcbRelayID2
    Dcombos.GetEmployees Me.DcbSuperID
    Dcombos.GetTblLocations Me.DcbLocatioID
    Dcombos.GetTblLocations Me.DcbLocatioID2
    Dcombos.GetTblTrips Me.DcbTypeTrip
    Dcombos.GetTblProgrammTypes Me.DcbProgramID
    Dcombos.GetTblTypePath Me.DcbTypePath
    If SystemOptions.UserInterface = ArabicInterface Then
    CBoBasedON.Clear
    CBoBasedON.AddItem "»·« "
    CBoBasedON.AddItem " «ﬂÌœ ÕÃ“  "
    Else
    CBoBasedON.Clear
    CBoBasedON.AddItem "NA"
    CBoBasedON.AddItem "Order"
    End If
        
        
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
    RsSavRec.Fields("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("RecordDateH").value = RecordDateH.value
    RsSavRec.Fields("BranchID").value = val(Me.Dcbranch.BoundText)
    '''////
        If txtNoteSerial1.Text = "" Then
              txtNoteSerial1.Text = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 72, 72, , , , , , val(SeasonsID.BoundText))
        End If
        RsSavRec.Fields("NoteSerial1").value = IIf(Me.txtNoteSerial1 <> "", val(txtNoteSerial1.Text), Null)
        
    RsSavRec.Fields("SeasonsID").value = val(SeasonsID.BoundText)
    RsSavRec.Fields("NoteSerialOrder").value = val(TxtNoteSerialOrder.Text)
    RsSavRec.Fields("EqupID").value = val(DcbEqupID.BoundText)
    RsSavRec.Fields("DayName1").value = lbl(11).Caption
    RsSavRec.Fields("DayName2").value = lbl(10).Caption
    RsSavRec.Fields("CurrDate").value = CurrDate.value
    RsSavRec.Fields("CurrDateH").value = CurrDateH.value
    RsSavRec.Fields("DriverName").value = TxtDriverName.Text
    RsSavRec.Fields("LocatioID").value = val(Me.DcbLocatioID.BoundText)
    RsSavRec.Fields("DriverID").value = val(DcbDriverID.BoundText)
    RsSavRec.Fields("HajzNo").value = val(TxtHajzNo.Text)
    RsSavRec.Fields("DiffTime").value = (TxtDiffTime.Text)
    RsSavRec.Fields("DiffDate").value = val(TxtDiffDate.Text)
    RsSavRec.Fields("basedOn").value = val(CBoBasedON.ListIndex)
    
 ''''//////////////////////
    RsSavRec.Fields("SuperVisName").value = TxtSuperVisName.Text
    RsSavRec.Fields("DriverName").value = TxtDriverName.Text
    RsSavRec.Fields("TypePath").value = val(Me.DcbTypePath.BoundText)
    RsSavRec.Fields("TypeTrip").value = val(DcbTypeTrip.BoundText)
    RsSavRec.Fields("ProgramID").value = val(Me.DcbProgramID.BoundText)
    RsSavRec.Fields("FromCity").value = val(DcbFromCity.BoundText)
    RsSavRec.Fields("ToCity").value = val(DcbToCity.BoundText)
    RsSavRec.Fields("TimeOut").value = FormatDateTime(Me.TimeOut.value, vbShortTime)
    RsSavRec.Fields("TimeIn").value = FormatDateTime(Me.TimeIn.value, vbShortTime)
    RsSavRec.Fields("RelayID").value = val(RelayID.BoundText)
    RsSavRec.Fields("RelayName").value = TxtRelayName.Text
    RsSavRec.Fields("Remarks").value = TxtRemarks.Text
    RsSavRec.Fields("Address").value = TxtAddress.Text
    RsSavRec.Fields("CurrDate2").value = CurrDate2.value
    RsSavRec.Fields("Phone").value = TxtPhone.Text
    RsSavRec.Fields("CurrDateH2").value = CurrDateH2.value
    RsSavRec.Fields("LocatioID2").value = val(DcbLocatioID2.BoundText)
    RsSavRec.Fields("RelayID2").value = val(DcbRelayID2.BoundText)
    RsSavRec.Fields("RelayName2").value = TxtRelayName2.Text
    RsSavRec.Fields("SuperID").value = val(DcbSuperID.BoundText)
If ChTypeDrive(1).value = True Then
RsSavRec.Fields("TypeDrive").value = 1
Else
RsSavRec.Fields("TypeDrive").value = 0
End If
If ChRelayType(1).value = True Then
RsSavRec.Fields("RelayType").value = 1
Else
RsSavRec.Fields("RelayType").value = 0
End If
If ChRelayType2(1).value = True Then
RsSavRec.Fields("RelayType2").value = 1
Else
RsSavRec.Fields("RelayType2").value = 0
End If
RsSavRec.Fields("PathID").value = val(DcbPath.BoundText)
RsSavRec.Fields("OrderID").value = val(TxtOrderID.Text)

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ''/////
    RsSavRec.Fields("UserID").value = IIf(DCboUserName.BoundText <> "", Trim(DCboUserName.BoundText), Null)
    RsSavRec.update

  
      Select Case Me.TxtModFlg.Text
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
    Dim i As Integer
loadcombo2
    Dim ContactTime As Date
    
    TxtPrintCount.Text = IIf(IsNull(RsSavRec.Fields("PrintCount").value), 0, RsSavRec.Fields("PrintCount").value)
    lbl(11).Caption = IIf(IsNull(RsSavRec.Fields("DayName1").value), "", RsSavRec.Fields("DayName1").value)
     Me.txtNoteSerial1.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial1").value), "", RsSavRec.Fields("NoteSerial1").value)
     lbl(10).Caption = IIf(IsNull(RsSavRec.Fields("DayName2").value), "", RsSavRec.Fields("DayName2").value)
    TxtSerial1.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TxtNoteSerialOrder.Text = IIf(IsNull(RsSavRec.Fields("NoteSerialOrder").value), "", RsSavRec.Fields("NoteSerialOrder").value)
    SeasonsID.BoundText = IIf(IsNull(RsSavRec.Fields("SeasonsID").value), "", RsSavRec.Fields("SeasonsID").value)
    XPDtbTrans.value = IIf(IsNull(RsSavRec.Fields("RecordDate").value), Date, RsSavRec.Fields("RecordDate").value)
    RecordDateH.value = IIf(IsNull(RsSavRec.Fields("RecordDateH").value), ToHijriDate(Date), RsSavRec.Fields("RecordDateH").value)
    Dcbranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchID").value), "", RsSavRec.Fields("BranchID").value)
    DCboUserName.BoundText = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    CurrDate.value = IIf(IsNull(RsSavRec.Fields("CurrDate").value), Date, RsSavRec.Fields("CurrDate").value)
    CurrDateH.value = IIf(IsNull(RsSavRec.Fields("CurrDateH").value), ToHijriDate(Date), RsSavRec.Fields("CurrDateH").value)
    Me.DcbLocatioID.BoundText = IIf(IsNull(RsSavRec.Fields("LocatioID").value), "", RsSavRec.Fields("LocatioID").value)
    TxtHajzNo.Text = IIf(IsNull(RsSavRec.Fields("HajzNo").value), "", RsSavRec.Fields("HajzNo").value)
    TxtDriverName.Text = IIf(IsNull(RsSavRec.Fields("DriverName").value), "", RsSavRec.Fields("DriverName").value)
    Me.DcbTypeTrip.BoundText = IIf(IsNull(RsSavRec.Fields("TypeTrip").value), 0, RsSavRec.Fields("TypeTrip").value)
    DcbProgramID.BoundText = IIf(IsNull(RsSavRec.Fields("ProgramID").value), "", RsSavRec.Fields("ProgramID").value)
    TxtSuperVisName.Text = IIf(IsNull(RsSavRec.Fields("SuperVisName").value), "", RsSavRec.Fields("SuperVisName").value)
    Me.DcbTypePath.BoundText = IIf(IsNull(RsSavRec.Fields("TypePath").value), 0, RsSavRec.Fields("TypePath").value)
    Me.RelayID.BoundText = IIf(IsNull(RsSavRec.Fields("RelayID").value), "", RsSavRec.Fields("RelayID").value)
    TxtRelayName.Text = IIf(IsNull(RsSavRec.Fields("RelayName").value), "", RsSavRec.Fields("RelayName").value)
    Me.DcbPath.BoundText = IIf(IsNull(RsSavRec.Fields("PathID").value), 0, RsSavRec.Fields("PathID").value)
    Me.TxtOrderID.Text = IIf(IsNull(RsSavRec.Fields("OrderID").value), 0, RsSavRec.Fields("OrderID").value)
    Me.TxtDiffTime.Text = IIf(IsNull(RsSavRec.Fields("DiffTime").value), 0, RsSavRec.Fields("DiffTime").value)
    Me.TxtDiffDate.Text = IIf(IsNull(RsSavRec.Fields("DiffDate").value), 0, RsSavRec.Fields("DiffDate").value)
    CBoBasedON.ListIndex = IIf(IsNull(RsSavRec.Fields("BasedOn").value), -1, RsSavRec.Fields("BasedOn").value)
    
    If Not IsNull(RsSavRec.Fields("TimeOut").value) Then
      ContactTime = FormatDateTime(RsSavRec.Fields("TimeOut").value, vbShortTime)
      Me.TimeOut.value = ContactTime
    End If
    If Not IsNull(RsSavRec.Fields("TimeIn").value) Then
      ContactTime = FormatDateTime(RsSavRec.Fields("TimeIn").value, vbShortTime)
      Me.TimeIn.value = ContactTime
    End If
    TxtRemarks.Text = IIf(IsNull(RsSavRec.Fields("Remarks").value), "", RsSavRec.Fields("Remarks").value)
    TxtAddress.Text = IIf(IsNull(RsSavRec.Fields("Address").value), "", RsSavRec.Fields("Address").value)
    CurrDate2.value = IIf(IsNull(RsSavRec.Fields("CurrDate2").value), Date, RsSavRec.Fields("CurrDate2").value)
    CurrDateH2.value = IIf(IsNull(RsSavRec.Fields("CurrDateH2").value), ToHijriDate(Date), RsSavRec.Fields("CurrDateH2").value)
    
    DcbLocatioID2.BoundText = IIf(IsNull(RsSavRec.Fields("LocatioID2").value), "", RsSavRec.Fields("LocatioID2").value)
    DcbRelayID2.BoundText = IIf(IsNull(RsSavRec.Fields("RelayID2").value), "", RsSavRec.Fields("RelayID2").value)
    Me.TxtRelayName2.Text = IIf(IsNull(RsSavRec.Fields("RelayName2").value), "", RsSavRec.Fields("RelayName2").value)
    DcbSuperID.BoundText = IIf(IsNull(RsSavRec.Fields("SuperID").value), "", RsSavRec.Fields("SuperID").value)
    TxtPhone.Text = IIf(IsNull(RsSavRec.Fields("Phone").value), "", RsSavRec.Fields("Phone").value)
    If Not (IsNull(RsSavRec.Fields("TypeDrive").value)) Then
    If RsSavRec.Fields("TypeDrive").value = 1 Then
    ChTypeDrive(1).value = True
    Else
    ChTypeDrive(0).value = True
    End If
    Else
    ChTypeDrive(0).value = True
    End If
    
    If Not (IsNull(RsSavRec.Fields("RelayType2").value)) Then
    If RsSavRec.Fields("RelayType2").value = 1 Then
    ChRelayType2(1).value = True
    Else
    ChRelayType2(0).value = True
    End If
    Else
    ChRelayType2(0).value = True
    End If
    If Not (IsNull(RsSavRec.Fields("RelayType").value)) Then
    If RsSavRec.Fields("RelayType").value = 1 Then
    ChRelayType(1).value = True
    Else
    ChRelayType(0).value = True
    End If
    Else
    ChRelayType(0).value = True
    End If
    Me.DcbDriverID.BoundText = IIf(IsNull(RsSavRec.Fields("DriverID").value), "", RsSavRec.Fields("DriverID").value)
    DcbEqupID.BoundText = IIf(IsNull(RsSavRec.Fields("EqupID").value), 0, RsSavRec.Fields("EqupID").value)
    CurrDate_Change
    CurrDate2_Change

    ''//////////
     LabCurrRec.Caption = RsSavRec.AbsolutePosition ': ProgressBar1.value = 50
     LabCountRec.Caption = RsSavRec.RecordCount ': ProgressBar1.value = 60

ErrTrap:
End Sub
Private Sub ISButton5_Click()
print_report
End Sub
' check before rece
'++++++++++++++++++++++++++++++++++++++++++++
Sub GetEmp(Optional EmpID As Double = 0)
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim Sql As String
Sql = "select * from TblEmployee where Emp_ID=" & EmpID & ""
Rs7.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs7.RecordCount > 0 Then
TxtPhone.Text = IIf(IsNull(Rs7("Emp_Phone").value), "", Rs7("Emp_Phone").value)
TxtAddress.Text = IIf(IsNull(Rs7("kafeladd").value), "", Rs7("kafeladd").value)
Else
TxtAddress.Text = ""
TxtPhone.Text = ""
End If
End Sub

Private Sub btnSave_Click()
   ' On Error GoTo ErrTrap
         Dim total As Double
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    Dim Sm As Double

    '---------------------- check if data Vaclete -----------------------
      If Dcbranch.Text = "" And val(Dcbranch.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·›—⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Branch ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            Dcbranch.SetFocus
            Exit Sub
     End If
     If val(Me.CBoBasedON.ListIndex) = 1 Then
        If val(TxtOrderID.Text) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· —ﬁ„  «ﬂÌœ «·ÕÃ“ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Eneter No ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            TxtOrderID.SetFocus
            Exit Sub
     End If
     End If
        If DcbEqupID.Text = "" And val(DcbEqupID.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·Õ«›·…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Vehicle ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbEqupID.SetFocus
            Exit Sub
     End If
        If ChTypeDrive(0).value = True Then
       If DcbDriverID.Text = "" And val(DcbDriverID.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«—  ﬁ«∆œ «·Õ«›·…", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Leader of Vehicle ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbDriverID.SetFocus
            Exit Sub
     End If
   Else
          If TxtDriverName.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «”„ ﬁ«∆œ «·Õ«›·… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Eneter Leader of Vehicle ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            TxtDriverName.SetFocus
            Exit Sub
     End If
    End If

     '    If ChRelayType(0).value = True Then
     '  If RelayID.text = "" And val(RelayID.BoundText) = 0 Then
   '     If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·„ÊŸ› «·„—Õ·", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        Else
     '       MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '     End If
     '       RelayID.SetFocus
    '        Exit Sub
     'End If
'   Else
    '      If TxtRelayName.text = "" Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
      '      MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·„ÊŸ› «·„—Õ·", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
     '       Else
     '       MsgBox "Please Eneter Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    '     End If
     '       TxtRelayName.SetFocus
    '        Exit Sub
     'End If
    'End If
   ' If ChRelayType2(0).value = True Then
    '   If DcbRelayID2.text = "" And val(DcbRelayID2.BoundText) = 0 Then
   '     If SystemOptions.UserInterface = ArabicInterface Then
   '         MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·„ÊŸ› «·„—Õ·", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
   '         Else
   '         MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
   '      End If
   '         DcbRelayID2.SetFocus
   '         Exit Sub
  '   End If
 '  Else
  '        If TxtRelayName2.text = "" Then
  '      If SystemOptions.UserInterface = ArabicInterface Then
  '          MsgBox "⁄›Ê« ...«·—Ã«¡ «œŒ«· «·„ÊŸ› «·„—Õ·", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
  '          Else
  ''          MsgBox "Please Eneter Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
  '       End If
  '          TxtRelayName2.SetFocus
  '          Exit Sub
  '   End If
     
 '  End If
  ' If Me.ChRelayType2(0).value = True Then
  '        If DcbRelayID2.text = "" And val(DcbRelayID2.BoundText) = 0 Then
  '      If SystemOptions.UserInterface = ArabicInterface Then
 '           MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— «·„ÊŸ› «·„—Õ·", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
 '           Else
 '           MsgBox "Please Select Employee ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 '        End If
 '           DcbRelayID2.SetFocus
 '           Exit Sub
  '   End If
  ' End If
     
    '   If DcbFromCity.text = "" And val(DcbFromCity.BoundText) = 0 Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ÃÂ… «·—Õ·… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
    '        Else
     ''       MsgBox "Please Select Flight destination ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
     '    End If
    '        DcbFromCity.SetFocus
    '        Exit Sub
   '  End If
   If DcbPath.Text = "" And val(DcbPath.BoundText) = 0 Then
       If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "⁄›Ê« ...«·—Ã«¡ ≈Œ Ì«— ÃÂ… «·—Õ·… ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Please Select Flight destination ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
         End If
            DcbPath.SetFocus
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
   Dim TxtNoteSerial1str As String

    If txtNoteSerial1.Text = "" Then
     TxtNoteSerial1str = Voucher_coding(val(Me.Dcbranch.BoundText), XPDtbTrans.value, 72, 72, , , , , , val(SeasonsID.BoundText))
                If TxtNoteSerial1str = "error" Then
                    MsgBox " ·« Ì„ﬂ‰ «÷«›…  Õ—ﬂ…  ÃœÌœ…  ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                Else
                               
                    If TxtNoteSerial1str = "" Then
                        MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  «·Õ—ﬂ… ÃœÌœ     ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                    End If
                End If
    End If
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
    StrRecID = new_id("TblDeported", "ID", "")
    RsSavRec.AddNew
    TxtSerial1.Text = StrRecID
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Private Sub RecordDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
      VBA.Calendar = vbCalGreg
    XPDtbTrans.value = ToGregorianDate(RecordDateH.value)
    End If
End Sub

Private Sub RelayID_Click(Area As Integer)
RelayID_Change
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.Text, EmpID
        DcbRelayID2.BoundText = EmpID
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text3.Text, EmpID
        DcbSuperID.BoundText = EmpID
    End If
End Sub

Private Sub TimeIn_Change()
ClCulteDiffDateTime
End Sub

Private Sub TxtHajzNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If Me.TxtModFlg.Text <> "R" Then
If val(TxtHajzNo.Text) <> 0 Then
TxtOrderID.Text = GetIDOrder(val(TxtNoteSerialOrder.Text), val(SeasonsID.BoundText))
TxtOrderID.Text = GetOrderID(val(TxtHajzNo.Text))
Ceckrep
End If
End If
End If
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtHajzNo.Text, 0)

End Sub

Public Sub Booking()
Dim EmpName  As String
Dim EmpMbile As String
Dim HajzNo As String
Dim SeasonsID1 As Double
Dim NoteSerial1 As Double
Ceckrep
loadcombo val(TxtOrderID.Text)
 GetOrdersData val(TxtOrderID), EmpName, EmpMbile, HajzNo, NoteSerial1, SeasonsID1
'TxtHajzNo.Text = HajzNo
If Me.TxtModFlg <> "R" Then
Me.TxtNoteSerialOrder.Text = NoteSerial1
Me.TxtSuperVisName.Text = EmpName
Me.TxtPhone.Text = EmpMbile
SeasonsID.BoundText = SeasonsID1
End If
End Sub
Private Sub TxtNoteSerialOrder_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If val(SeasonsID.BoundText) <> 0 Then
TxtOrderID.Text = GetIDOrder(val(TxtNoteSerialOrder.Text), val(SeasonsID.BoundText))
Else
MsgBox "Ì—ÃÏ «Œ Ì«— «·„Ê”„"
Exit Sub
End If
Booking
End If
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoteSerialOrder.Text, 0)
End Sub

Private Sub TxtNoteSerialOrder_KeyUp(KeyCode As Integer, Shift As Integer)
If Me.TxtModFlg.Text <> "R" Then
If KeyCode = vbKeyF3 Then
    Unload FrmSearch_Hajj
         FrmSearch_Hajj.SendForm = "BookingRequest12"
         FrmSearch_Hajj.show
End If
End If
End Sub

 Public Function GetOrderID(Optional ByRef HajzNo As Double) As Double
Dim Rs4 As ADODB.Recordset
Dim Sql As String
Set Rs4 = New ADODB.Recordset
Sql = "SELECT    id from tblbookingrequest2 where OrdeNo=" & HajzNo
 
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetOrderID = IIf(IsNull(Rs4("ID").value), 0, Rs4("ID").value)
Else
GetOrderID = 0
End If
End Function
Function GetNoVehicleNo(Optional ID As Double) As Double
Dim Rs4 As ADODB.Recordset
Dim Sql As String
Set Rs4 = New ADODB.Recordset
Sql = "SELECT     COUNT(OrderID) AS CuntOrder"
Sql = Sql & " From dbo.TblDeported"
Sql = Sql & " WHERE     (OrderID  = " & ID & ")"
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetNoVehicleNo = IIf(IsNull(Rs4("CuntOrder").value), 0, Rs4("CuntOrder").value)
Else
GetNoVehicleNo = 0
End If
End Function
Function GetNoPath(Optional ID As Double) As Double
Dim Rs4 As ADODB.Recordset
Dim Sql As String
Set Rs4 = New ADODB.Recordset
Sql = " SELECT     COUNT(HID) AS CountPath, HID"
Sql = Sql & " From dbo.TblFlightDetails"
Sql = Sql & " Where (HID = " & ID & ")"
Sql = Sql & " GROUP BY HID"
Rs4.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
GetNoPath = IIf(IsNull(Rs4("CountPath").value), 0, Rs4("CountPath").value)
Else
GetNoPath = 0
End If
End Function
Sub loadcombo2(Optional ID As Double)
Dim str As String
If SystemOptions.UserInterface = ArabicInterface Then
  str = " SELECT     ID, Name"
  str = str & " FROM         dbo.TblShrines"
 Else
 str = " SELECT     ID, NameE"
 str = str & " FROM         dbo.TblShrines"
 End If
   fill_combo DcbPath, str
End Sub

Sub loadcombo(Optional ID As Double)
Dim str As String
If SystemOptions.UserInterface = ArabicInterface Then
str = " SELECT     ID, Name"
  str = str & " FROM         dbo.TblShrines"
  str = str & "   WHERE     (ID IN"
  str = str & "                         (SELECT     PathID"
  str = str & "                            From TblFlightDetails"
   str = str & "    WHERE     HID = " & ID & "))"
 Else
 str = " SELECT     ID, NameE"
  str = str & " FROM         dbo.TblShrines"
  str = str & "   WHERE     (ID IN"
  str = str & "                         (SELECT     PathID"
  str = str & "                            From TblFlightDetails"
  str = str & "    WHERE     HID = " & ID & "))"
 
 End If
   fill_combo DcbPath, str
End Sub





Public Sub TxtOrderID_KeyPress(KeyAscii As Integer)

End Sub
Sub Ceckrep()
Dim VehiclNo As Double
Dim ProgramID As Double
Dim VehicleNo As Double
If Me.TxtModFlg.Text <> "R" Then
If val(TxtOrderID.Text) <> 0 Then
RetriveOrderInformation val(TxtOrderID.Text), ProgramID, VehicleNo
VehicleNo = VehicleNo * GetNoPath(val(TxtOrderID.Text))
VehiclNo = GetNoVehicleNo(val(TxtOrderID.Text))
If Round(VehiclNo, 2) >= Round(VehicleNo, 2) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " ·«ÌÊÃœ Õ«›·«  „ «Õ… ·Â–« «·«„—"
Else
MsgBox "There are no Vehicle available"
End If
RetriveOrderInformation 0, ProgramID, VehicleNo
loadcombo 0
TxtOrderID.Text = 0
Exit Sub
End If
DcbProgramID.BoundText = ProgramID
End If
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcbDriverID.BoundText = EmpID
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
    Dim Sql As String
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
    If Me.TxtModFlg.Text <> "R" Then
        Select Case Me.TxtModFlg.Text
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
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    If TxtSerial1.Text <> "" Then
        TxtModFlg = "E"
            ClCulteDiffDateTime
        Me.DCboUserName.BoundText = user_id
        Frm2.Enabled = True
        Me.Dcbranch.SetFocus
          CurrDate2_Change
  XPDtbTrans_Change
  CurrDate_Change
  loadcombo val(TxtOrderID.Text)
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
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
     clear_all Me
     Dcbranch.BoundText = Current_branch
     ChRelayType_Click (0)
     ChRelayType(0).value = True
     ChTypeDrive(0).value = True
     ChRelayType2(0).value = True
     ChRelayType2_Click (0)
     ChTypeDrive_Click (0)

    TxtModFlg.Text = "N"
    CBoBasedON.ListIndex = 1
CBoBasedON_Change

    Me.DCboUserName.BoundText = user_id
    Me.Dcbranch.BoundText = Current_branch
    Dcbranch.SetFocus
  TimeOut.value = Time
  TimeIn.value = Time
  Me.SeasonsID.BoundText = GetMosim(0)
  CurrDate2_Change
  XPDtbTrans_Change
  CurrDate_Change
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
Function print_report(Optional NoteSerial As String, Optional reportno As Integer = 0)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
  MySQL = " SELECT     dbo.TblDeported.ID, dbo.TblDeported.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblDeported.RecordDate, "
  MySQL = MySQL & "                    dbo.TblDeported.RecordDateH, dbo.TblDeported.CurrDate, dbo.TblDeported.CurrDateH, dbo.TblDeported.DriverName, dbo.TblDeported.TimeOut,"
  MySQL = MySQL & "                    dbo.TblDeported.TimeIn, dbo.TblDeported.Remarks, dbo.TblDeported.Address, dbo.TblDeported.CurrDate2, dbo.TblDeported.CurrDateH2, dbo.TblDeported.Phone,"
  MySQL = MySQL & "                    dbo.TblDeported.DayName1, dbo.TblDeported.DayName2, dbo.TblDeported.TypeDrive, dbo.TblDeported.TypeTrip, dbo.TblTrips.Name AS TripName,"
  MySQL = MySQL & "                    dbo.TblTrips.NameE AS TripNameE, dbo.TblDeported.ProgramID, dbo.TblProgrammTypes.Name AS ProgName, dbo.TblProgrammTypes.NameE AS ProgNameE,"
  MySQL = MySQL & "                    dbo.TblDeported.EqupID, dbo.TblCarsData.BoardNO, dbo.TblDeported.OrderID, dbo.TblDeported.UserID, dbo.TblUsers.UserName, dbo.TblDeported.SuperID,"
  MySQL = MySQL & "                    dbo.TblDeported.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE, dbo.TblCarsData.OperatorN, dbo.TblDeported.SuperVisName,"
  MySQL = MySQL & "                    dbo.TblDeported.HajzNo, dbo.TblCarsData.Fullcode, dbo.TblCarsData.Name, dbo.TblDeported.TypePath, dbo.TblHotels.Name AS TypPathName,"
  MySQL = MySQL & "                    dbo.TblHotels.NameE AS TypPathNameE, dbo.TblDeported.DriverID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode AS DriverFullcode,"
  MySQL = MySQL & "                    dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
  MySQL = MySQL & "                    dbo.TblDeported.LocatioID, TblLocations_1.Name AS LoctionName, TblLocations_1.NameE AS LoctionNameE, dbo.TblDeported.LocatioID2,"
  MySQL = MySQL & "                    TblLocations_1.Name AS LoctionName2, TblLocations_1.NameE AS LoctionNameE2, dbo.TblDeported.BasedOn, dbo.TblDeported.DiffDate, dbo.TblDeported.DiffTime,"
  MySQL = MySQL & "                    dbo.TblDeported.SeasonsID , dbo.TblDeported.NoteSerialOrder, dbo.TblDeported.NoteSerial1 , dbo.TblDeported.Prefix, dbo.TblDeported.PrintCount"
  MySQL = MySQL & " FROM         dbo.TblDeported LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblLocations TblLocations_1 ON dbo.TblDeported.LocatioID2 = TblLocations_1.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblLocations TblLocations_2 ON dbo.TblDeported.LocatioID = TblLocations_2.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblEmployee ON dbo.TblDeported.DriverID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  MySQL = MySQL & "                     dbo.TblHotels ON dbo.TblDeported.TypePath = dbo.TblHotels.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblShrines ON dbo.TblDeported.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblUsers ON dbo.TblDeported.UserID = dbo.TblUsers.UserID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblCarsData ON dbo.TblDeported.EqupID = dbo.TblCarsData.id LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblProgrammTypes ON dbo.TblDeported.ProgramID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblTrips ON dbo.TblDeported.TypeTrip = dbo.TblTrips.ID LEFT OUTER JOIN"
  MySQL = MySQL & "                    dbo.TblBranchesData ON dbo.TblDeported.BranchID = dbo.TblBranchesData.branch_id"
  MySQL = MySQL & "  Where (dbo.TblDeported.id =" & val(TxtSerial1.Text) & ")"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDeportedPilgrims.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDeportedPilgrims.rpt"
        End If
        If reportno = 0 Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDeportedPilgrims.rpt"
        ElseIf reportno = 1 Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDeportedPilgrims1.rpt"
         
         ElseIf reportno = 2 Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDeportedPilgrims2.rpt"
         End If
         
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
  If reportno = 0 Then
 Cn.Execute "Update TblDeported set PrintCount=" & val(TxtPrintCount.Text) & "+1  where id=" & val(TxtSerial1.Text) & ""
 RsSavRec.Requery
 FindRec val(TxtSerial1.Text)
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
 xReport.ParameterFields(12).AddCurrentValue TxtHajzNo.Text
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
         xReport.ParameterFields(12).AddCurrentValue TxtHajzNo.Text
         
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
    Me.Caption = "Transport Hajj "
    Me.lbl(4).Caption = "No"
    Me.lbl(2).Caption = "Date"
    Me.lbl(7).Caption = "Branch"
    Me.Label1(2).Caption = Me.Caption
    Frame7.Caption = "Leader Vehicle "
    lbl(3).Caption = "Vehicle No."
    lbl(1).Caption = "Trip Type"
    lbl(27).Caption = "Time Out"
    lbl(23).Caption = "Trip"
    lbl(24).Caption = "From"
    lbl(25).Caption = "To"
    lbl(17).Caption = "Date"
    lbl(6).Caption = "Time Arrival"
    lbl(18).Caption = "Arrival Data"
    lbl(13).Caption = "Program Type"
    ChTypeDrive(0).RightToLeft = False
    ChTypeDrive(0).Caption = "Employee"
    ChTypeDrive(1).RightToLeft = False
    ChTypeDrive(1).Caption = "Nan"
    lbl(0).Caption = "Date"
  lbl(9).Caption = "Remarks"
 lbl(16).Caption = "Phone"
lbl(5).Caption = "Location"
lbl(12).Caption = "Location"
Frame8.Caption = "Relay"
ChRelayType(0).RightToLeft = False
ChRelayType(0).Caption = "Employee"
ChRelayType(1).RightToLeft = False
ChRelayType(1).Caption = "Nan"
Frame13.Caption = "Relay"
ChRelayType2(0).RightToLeft = False
ChRelayType2(0).Caption = "Employee"
ChRelayType2(1).RightToLeft = False
ChRelayType2(1).Caption = "Nan"
Frame14.Caption = "Supervisor Data"
lbl(14).Caption = "Supervisor"
lbl(15).Caption = "Address"
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

ErrTrap:
End Sub
Private Sub AddNewRecored()
   Dim My_SQL As String
   Dim rs As ADODB.Recordset
  On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   My_SQL = "TblDeported"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial1.Text = rs.RecordCount + 1
    Else
        TxtSerial1.Text = 1
    End If
   rs.Close
ErrTrap:
End Sub
'+++++++++++++++++++++++++++++++++ en

Private Sub TxtSwapCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSwapCode.Text, EmpID
        RelayID.BoundText = EmpID
    End If
End Sub



Private Sub XPDtbTrans_Change()
If Me.TxtModFlg.Text <> "R" Then
         RecordDateH.value = ToHijriDate(XPDtbTrans.value)
End If
End Sub
