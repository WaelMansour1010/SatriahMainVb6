VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmProductionPlan 
   BackColor       =   &H00E2E9E9&
   Caption         =   " "
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   HelpContextID   =   580
   Icon            =   "FrmProductionPlan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   18960
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   10950
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   18960
      _cx             =   33443
      _cy             =   19315
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
      BackColor       =   12648447
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   7
      BorderWidth     =   2
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   5025
         Left            =   30
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   -90
         Width           =   18870
         _cx             =   33285
         _cy             =   8864
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
         Begin VB.Frame FrameBranch 
            Caption         =   "Õœœ «·›—Ê⁄"
            Height          =   3735
            Left            =   0
            TabIndex        =   82
            Top             =   1260
            Width           =   9135
            Begin VB.ListBox ListGroupAll 
               Height          =   1620
               ItemData        =   "FrmProductionPlan.frx":038A
               Left            =   4800
               List            =   "FrmProductionPlan.frx":0391
               TabIndex        =   84
               Top             =   360
               Width           =   4215
            End
            Begin VB.ListBox ListGroupSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1620
               ItemData        =   "FrmProductionPlan.frx":03A3
               Left            =   120
               List            =   "FrmProductionPlan.frx":03AA
               TabIndex        =   83
               Top             =   360
               Width           =   4215
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   9
               Left            =   240
               TabIndex        =   90
               Top             =   2160
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmProductionPlan.frx":03C1
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lblPlantype 
               Height          =   15
               Left            =   3960
               TabIndex        =   117
               Top             =   2400
               Visible         =   0   'False
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
               Left            =   4320
               TabIndex        =   89
               Top             =   600
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
               Left            =   4320
               TabIndex        =   88
               Top             =   840
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
               Left            =   4320
               TabIndex        =   87
               Top             =   1080
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
               Left            =   4320
               TabIndex        =   86
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               Caption         =   "«·›—Ê⁄ «·„Õœœ…"
               Height          =   255
               Left            =   1320
               TabIndex        =   85
               Top             =   120
               Width           =   1335
            End
         End
         Begin VB.ComboBox CplanType 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmProductionPlan.frx":075B
            Left            =   9960
            List            =   "FrmProductionPlan.frx":075D
            TabIndex        =   70
            Top             =   825
            Width           =   4695
         End
         Begin VB.Frame Frame2 
            Caption         =   " »‰«¡ ⁄·Ï"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3435
            Left            =   9240
            TabIndex        =   62
            Top             =   1545
            Width           =   9615
            Begin VB.TextBox TxtNoteserial 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2880
               TabIndex        =   106
               Top             =   1680
               Width           =   1695
            End
            Begin VB.ComboBox CBoBaseSanad 
               Height          =   315
               ItemData        =   "FrmProductionPlan.frx":075F
               Left            =   5580
               List            =   "FrmProductionPlan.frx":0761
               Style           =   2  'Dropdown List
               TabIndex        =   105
               Top             =   1680
               Width           =   2550
            End
            Begin VB.TextBox TxtStoreID 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   7245
               TabIndex        =   102
               Top             =   1320
               Width           =   885
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   7200
               TabIndex        =   101
               Top             =   600
               Width           =   930
            End
            Begin VB.TextBox TxtCashCustomerName 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2910
               TabIndex        =   100
               Top             =   960
               Width           =   5235
            End
            Begin VB.TextBox txtRemarks 
               Alignment       =   1  'Right Justify
               Height          =   432
               Left            =   240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   97
               Top             =   2040
               Width           =   7890
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ·»«  œ«Œ·Ì…"
               Height          =   255
               Index           =   4
               Left            =   3120
               TabIndex        =   92
               Top             =   240
               Width           =   1692
            End
            Begin VB.TextBox TxtOldPlanNo 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   81
               Top             =   240
               Width           =   732
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ·»«  ‘Õ‰  "
               Height          =   255
               Index           =   2
               Left            =   4800
               TabIndex        =   80
               Top             =   240
               Width           =   1812
            End
            Begin VB.Frame Frame3 
               Caption         =   "Õœœ «·› —… ·«” œ⁄«¡ «·»Ì«‰« "
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   120
               TabIndex        =   75
               Top             =   960
               Width           =   2055
               Begin MSComCtl2.DTPicker dbFromDate 
                  Height          =   276
                  Left            =   120
                  TabIndex        =   76
                  Top             =   240
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   476
                  _Version        =   393216
                  Format          =   94109697
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker DBTo 
                  Height          =   276
                  Left            =   120
                  TabIndex        =   77
                  Top             =   600
                  Width           =   1320
                  _ExtentX        =   2328
                  _ExtentY        =   476
                  _Version        =   393216
                  Format          =   94109697
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰"
                  Height          =   276
                  Index           =   5
                  Left            =   1476
                  TabIndex        =   79
                  Top             =   240
                  Width           =   468
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï"
                  Height          =   276
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   78
                  Top             =   600
                  Width           =   600
               End
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "ŒÿÂ ”«»ﬁ…"
               Height          =   255
               Index           =   3
               Left            =   1800
               TabIndex        =   65
               Top             =   240
               Width           =   1332
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "›Ê« Ì— „»Ì⁄« "
               Height          =   255
               Index           =   1
               Left            =   6600
               TabIndex        =   64
               Top             =   240
               Width           =   1572
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               Caption         =   "√„— »Ì⁄  "
               Height          =   255
               Index           =   0
               Left            =   8160
               TabIndex        =   63
               Top             =   240
               Width           =   1332
            End
            Begin MSDataListLib.DataCombo DcbCustomer 
               Height          =   315
               Left            =   2880
               TabIndex        =   103
               Top             =   600
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               BoundColumn     =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCboStoreName 
               Height          =   315
               Left            =   2880
               TabIndex        =   104
               Top             =   1320
               Width           =   4365
               _ExtentX        =   7699
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "7"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "Õœœ —ﬁ„ «·Œÿ…"
               Height          =   372
               Left            =   840
               TabIndex        =   109
               Top             =   240
               Width           =   852
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "·„” ‰œ „⁄Ì‰"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   17
               Left            =   8400
               TabIndex        =   108
               Top             =   1680
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·—ﬁ„"
               ForeColor       =   &H00000000&
               Height          =   372
               Index           =   16
               Left            =   4200
               TabIndex        =   107
               Top             =   1680
               Width           =   1056
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„·«ÕŸ« "
               Height          =   195
               Index           =   3
               Left            =   8640
               TabIndex        =   99
               Top             =   2160
               Width           =   705
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Height          =   315
               Index           =   9
               Left            =   8520
               TabIndex        =   98
               Top             =   2160
               Width           =   720
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„Œ“‰ «·„Õœœ"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   15
               Left            =   8400
               TabIndex        =   96
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì· «·„Õœœ"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   42
               Left            =   8370
               TabIndex        =   95
               Top             =   600
               Width           =   1050
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   8130
               TabIndex        =   94
               Top             =   1005
               Width           =   1290
            End
         End
         Begin VB.TextBox TxtTbllProductionPlanD 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   16140
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   825
            Width           =   1440
         End
         Begin VB.CheckBox ChkLocked 
            Alignment       =   1  'Right Justify
            Caption         =   "«Ìﬁ«› «· ⁄«„·"
            Height          =   168
            Left            =   18900
            TabIndex        =   31
            Top             =   915
            Visible         =   0   'False
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker dbTodate 
            Height          =   450
            Left            =   120
            TabIndex        =   33
            Top             =   825
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   794
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker DbFrom 
            Height          =   450
            Left            =   2040
            TabIndex        =   66
            Top             =   825
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   794
            _Version        =   393216
            Format          =   94109697
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   285
            Left            =   19080
            TabIndex        =   71
            Top             =   3075
            Visible         =   0   'False
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   4200
            TabIndex        =   73
            Top             =   825
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   556
            _Version        =   393216
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   735
            Index           =   5
            Left            =   0
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   0
            Width           =   18900
            _cx             =   33338
            _cy             =   1296
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
            Caption         =   " «· ŒÿÌÿ  "
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   0
            ChildSpacing    =   0
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   7
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
            Begin VB.TextBox txtnots2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   120
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   240
               Visible         =   0   'False
               Width           =   975
            End
            Begin ImpulseButton.ISButton XPBtnMove 
               Height          =   345
               Index           =   0
               Left            =   1860
               TabIndex        =   113
               Top             =   105
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   609
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
               ButtonImage     =   "FrmProductionPlan.frx":0763
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
               Height          =   345
               Index           =   3
               Left            =   1005
               TabIndex        =   114
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   609
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
               ButtonImage     =   "FrmProductionPlan.frx":0AFD
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
               Height          =   345
               Index           =   1
               Left            =   2670
               TabIndex        =   115
               Top             =   105
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   609
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
               ButtonImage     =   "FrmProductionPlan.frx":0E97
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
               Height          =   345
               Index           =   2
               Left            =   165
               TabIndex        =   116
               Top             =   105
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   609
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
               ButtonImage     =   "FrmProductionPlan.frx":1231
               ColorHighlight  =   4194304
               ColorHoverText  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
               ColorToggledHoverText=   16777215
               ColorTextShadow =   16777215
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰œÊ» «·„»Ì⁄« "
            Height          =   225
            Index           =   14
            Left            =   19200
            TabIndex        =   72
            Top             =   3075
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·ŒÿÂ"
            Height          =   435
            Index           =   13
            Left            =   14775
            TabIndex        =   69
            Top             =   825
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            Height          =   255
            Index           =   12
            Left            =   1440
            TabIndex        =   68
            Top             =   825
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   255
            Index           =   11
            Left            =   3480
            TabIndex        =   67
            Top             =   825
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Height          =   300
            Index           =   10
            Left            =   1560
            TabIndex        =   61
            Top             =   1260
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·Œÿ…"
            Height          =   195
            Index           =   7
            Left            =   17640
            TabIndex        =   35
            Top             =   825
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   435
            Index           =   0
            Left            =   9120
            TabIndex        =   34
            Top             =   825
            Width           =   585
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6465
         Left            =   30
         TabIndex        =   1
         Top             =   3270
         Width           =   18900
         _cx             =   33338
         _cy             =   11404
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   12648447
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·«’‰«›|«·„Ã„Ê⁄« "
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   1
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6045
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   18810
            _cx             =   33179
            _cy             =   10663
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
            Begin VB.TextBox TXTiTEMcODE 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   13905
               TabIndex        =   74
               Top             =   1905
               Width           =   1695
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ· «·«’‰«›"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   17145
               TabIndex        =   37
               Top             =   1965
               Width           =   1320
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ Ì«— ’‰›"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   15645
               TabIndex        =   36
               Top             =   1965
               Value           =   -1  'True
               Width           =   1530
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   3645
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   2340
               Width           =   18690
               _cx             =   32967
               _cy             =   6429
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
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3390
                  Left            =   120
                  TabIndex        =   60
                  Top             =   120
                  Width           =   18405
                  _cx             =   32464
                  _cy             =   5980
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
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   29
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmProductionPlan.frx":15CB
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
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   135
                  Left            =   13125
                  TabIndex        =   29
                  Top             =   2340
                  Width           =   1500
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   132
                  Left            =   6705
                  TabIndex        =   27
                  Text            =   "0"
                  Top             =   2535
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   204
                  Left            =   13125
                  TabIndex        =   18
                  Top             =   2400
                  Width           =   2655
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   120
                  Index           =   0
                  Left            =   -4800
                  TabIndex        =   8
                  Top             =   5985
                  Width           =   2610
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   7140
                  TabIndex        =   6
                  Top             =   2415
                  Visible         =   0   'False
                  Width           =   2580
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   13125
                  TabIndex        =   10
                  Top             =   990
                  Width           =   5145
                  _ExtentX        =   9075
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo dcproject 
                  Height          =   315
                  Left            =   14340
                  TabIndex        =   11
                  Top             =   780
                  Width           =   2040
                  _ExtentX        =   3598
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo Dcterm 
                  Height          =   315
                  Left            =   13275
                  TabIndex        =   26
                  Top             =   450
                  Width           =   3720
                  _ExtentX        =   6562
                  _ExtentY        =   556
                  _Version        =   393216
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
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   150
                  Index           =   8
                  Left            =   12120
                  TabIndex        =   9
                  Top             =   1320
                  Width           =   2220
               End
               Begin VB.Label Label55 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   60
                  Left            =   16830
                  TabIndex        =   7
                  Top             =   540
                  Width           =   1290
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Index           =   20
               Left            =   5505
               TabIndex        =   38
               Top             =   1860
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
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
               ButtonImage     =   "FrmProductionPlan.frx":19FD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Index           =   21
               Left            =   4680
               TabIndex        =   39
               Top             =   1860
               Width           =   795
               _ExtentX        =   1402
               _ExtentY        =   661
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
               ButtonImage     =   "FrmProductionPlan.frx":1D97
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo dcitems 
               Height          =   315
               Left            =   6465
               TabIndex        =   40
               Top             =   1905
               Width           =   7425
               _ExtentX        =   13097
               _ExtentY        =   556
               _Version        =   393216
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   375
               Index           =   11
               Left            =   120
               TabIndex        =   93
               Top             =   1965
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   661
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
               ButtonImage     =   "FrmProductionPlan.frx":2331
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   360
               Index           =   1
               Left            =   13785
               TabIndex        =   3
               Top             =   2310
               Visible         =   0   'False
               Width           =   1125
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6045
            Index           =   0
            Left            =   19545
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   45
            Width           =   18810
            _cx             =   33179
            _cy             =   10663
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
            Begin VB.OptionButton Option4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ Ì«— „Ã„Ê⁄Â"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   8820
               TabIndex        =   43
               Top             =   1200
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ ﬂ«›Â «·„Ã„Ê⁄« "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   10320
               TabIndex        =   42
               Top             =   1200
               Width           =   1800
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4035
               Index           =   3
               Left            =   0
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   1560
               Width           =   15345
               _cx             =   27067
               _cy             =   7117
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
               Begin VSFlex8Ctl.VSFlexGrid Grid1 
                  Height          =   2505
                  Left            =   0
                  TabIndex        =   59
                  Top             =   165
                  Width           =   12660
                  _cx             =   22331
                  _cy             =   4419
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
                  AllowUserResizing=   1
                  SelectionMode   =   1
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   1
                  Cols            =   21
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmProductionPlan.frx":28CB
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
               End
               Begin VB.TextBox Text2 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   5880
                  TabIndex        =   49
                  Top             =   2460
                  Visible         =   0   'False
                  Width           =   2175
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   1
                  Left            =   -3960
                  TabIndex        =   48
                  Top             =   6300
                  Width           =   2190
               End
               Begin VB.CheckBox Check3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   225
                  Left            =   10620
                  TabIndex        =   47
                  Top             =   2415
                  Width           =   2235
               End
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   5445
                  TabIndex        =   46
                  Text            =   "0"
                  Top             =   2580
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CheckBox Check2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   240
                  Left            =   10620
                  TabIndex        =   45
                  Top             =   2295
                  Width           =   1500
               End
               Begin MSDataListLib.DataCombo DataCombo1 
                  Height          =   315
                  Left            =   10620
                  TabIndex        =   50
                  Top             =   1335
                  Width           =   4305
                  _ExtentX        =   7594
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   315
                  Left            =   11850
                  TabIndex        =   51
                  Top             =   1080
                  Width           =   1620
                  _ExtentX        =   2858
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin MSDataListLib.DataCombo DataCombo3 
                  Height          =   315
                  Left            =   10770
                  TabIndex        =   52
                  Top             =   585
                  Width           =   3300
                  _ExtentX        =   5821
                  _ExtentY        =   556
                  _Version        =   393216
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
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Left            =   13905
                  TabIndex        =   54
                  Top             =   705
                  Width           =   870
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   180
                  Index           =   4
                  Left            =   10035
                  TabIndex        =   53
                  Top             =   1770
                  Width           =   1800
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   7
               Left            =   3705
               TabIndex        =   55
               Top             =   1080
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
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
               ButtonImage     =   "FrmProductionPlan.frx":2BF8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   8
               Left            =   3000
               TabIndex        =   56
               Top             =   1080
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   688
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
               ButtonImage     =   "FrmProductionPlan.frx":2F92
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DcGroup 
               Height          =   315
               Left            =   4665
               TabIndex        =   57
               Top             =   1200
               Width           =   4080
               _ExtentX        =   7197
               _ExtentY        =   556
               _Version        =   393216
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
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   6
               Left            =   13920
               TabIndex        =   58
               Top             =   810
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1185
         Left            =   30
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   9615
         Width           =   18900
         _cx             =   33338
         _cy             =   2090
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
         AutoSizeChildren=   7
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   285
            Left            =   11865
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   180
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14737632
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
            ButtonImage     =   "FrmProductionPlan.frx":352C
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   285
            Left            =   12750
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   180
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
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
            ButtonImage     =   "FrmProductionPlan.frx":38C6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   255
            Left            =   13950
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   225
            Visible         =   0   'False
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   450
            ButtonStyle     =   1
            ButtonPositionImage=   2
            Caption         =   ""
            BackColor       =   14871017
            FontSize        =   14.25
            FontName        =   "Arial"
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmProductionPlan.frx":3C60
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   0
            Left            =   10605
            TabIndex        =   19
            Top             =   435
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   1
            Left            =   9705
            TabIndex        =   20
            Top             =   435
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   2
            Left            =   8880
            TabIndex        =   21
            Top             =   435
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   390
            Index           =   3
            Left            =   7875
            TabIndex        =   22
            Top             =   435
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   4
            Left            =   6840
            TabIndex        =   23
            Top             =   435
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   390
            Index           =   6
            Left            =   4320
            TabIndex        =   24
            Top             =   435
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   5
            Left            =   5910
            TabIndex        =   25
            Top             =   435
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   330
            Left            =   17025
            TabIndex        =   28
            Tag             =   "Delete Row"
            Top             =   90
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   582
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmProductionPlan.frx":3FFA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   10
            Left            =   5280
            TabIndex        =   91
            Top             =   435
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
            ButtonStyle     =   1
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
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
            Height          =   195
            Left            =   1560
            TabIndex        =   17
            Top             =   180
            Width           =   1740
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
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
            Height          =   180
            Left            =   4920
            TabIndex        =   16
            Top             =   195
            Width           =   1515
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
      Top             =   6840
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "⁄—÷"
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
      ButtonImage     =   "FrmProductionPlan.frx":4016
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmProductionPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal x As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long
Function CuurentLogdata(Optional Currentmode As String)

    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·« ›«ﬁÌ…    " & TxtTbllProductionPlanD.Text & CHR(13) & " «·⁄„»· " & "" & CHR(13) & "  „œ Â« „‰  " & dbFromDate & CHR(13) & "  «·Ï " & dbTodate & CHR(13) & "  „·«ÕŸ«  " & TxtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & CHR(13) & "   „ «Ìﬁ«› «· ⁄«„· "
    End If
                    
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Contract No    " & TxtTbllProductionPlanD.Text & CHR(13) & " Customer " & "" & CHR(13) & " From   " & dbFromDate & CHR(13) & "  To  " & dbTodate & CHR(13) & "  Remarks " & TxtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & CHR(13) & " Locked "
    End If
                    
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
    
End Function

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
End Sub
Function print_report(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT     dbo.TbllProductionPlan.TbllProductionPlanD, dbo.TbllProductionPlan.planType, dbo.TbllProductionPlan.BranchId, TblBranchesData_3.branch_name, "
MySQL = MySQL & "                       TblBranchesData_3.branch_namee, dbo.TbllProductionPlan.Todate1, dbo.TbllProductionPlan.FromDate1, dbo.TbllProductionPlan.DbFrom,"
MySQL = MySQL & "                       dbo.TbllProductionPlan.DBTo, dbo.TbllProductionPlan.Opt, dbo.TbllProductionPlan.OldPlanNo, dbo.TbllProductionPlan.Remarks,"
MySQL = MySQL & "                       dbo.TblPlanBranches.BranchID AS PlanBranchID, TblBranchesData_1.branch_name AS Planbranch_name, TblBranchesData_1.branch_namee AS Planbranch_namee,"
MySQL = MySQL & "                       dbo.TbllProductionPlan.FromDate, dbo.TbllProductionPlan.Todate, dbo.TbllProductionPlan.Locked, dbo.TbllProductionPlanDetails.UnitID, dbo.TblUnites.UnitName,"
MySQL = MySQL & "                       dbo.TblUnites.UnitNamee, dbo.TbllProductionPlanDetails.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
MySQL = MySQL & "                       dbo.TbllProductionPlanDetails.Carid, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Name, dbo.TblCarsData.Model, TblBranchesData_2.branch_id,"
MySQL = MySQL & "                       TblBranchesData_2.branch_name AS branch_nameDet, TblBranchesData_2.branch_namee AS branch_nameeDet, dbo.TbllProductionPlanDetails.Driverid,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee1, dbo.TbllProductionPlanDetails.Price, dbo.TbllProductionPlanDetails.Discount, dbo.TbllProductionPlanDetails.Remarks AS RemarksD,"
MySQL = MySQL & "                       dbo.TbllProductionPlanDetails.Cunt, dbo.TbllProductionPlan.CustomerId, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
MySQL = MySQL & "                       dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TbllProductionPlan.NoteSerial, dbo.TbllProductionPlan.CashCustomerName, dbo.TbllProductionPlan.BaseSanad,"
MySQL = MySQL & "                       dbo.TbllProductionPlan.StoreId , dbo.TblStore.StoreName, dbo.TblStore.StoreNamee"
MySQL = MySQL & "  FROM         dbo.TblStore RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TbllProductionPlan ON dbo.TblStore.StoreID = dbo.TbllProductionPlan.StoreID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.TbllProductionPlan.CustomerId = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData TblBranchesData_2 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TbllProductionPlanDetails ON dbo.TblEmployee.Emp_ID = dbo.TbllProductionPlanDetails.Driverid ON"
MySQL = MySQL & "                       TblBranchesData_2.branch_id = dbo.TbllProductionPlanDetails.BranchId LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCarsData ON dbo.TbllProductionPlanDetails.Carid = dbo.TblCarsData.id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblItems ON dbo.TbllProductionPlanDetails.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblUnites ON dbo.TbllProductionPlanDetails.UnitID = dbo.TblUnites.UnitID ON"
MySQL = MySQL & "                       dbo.TbllProductionPlan.TbllProductionPlanD = dbo.TbllProductionPlanDetails.TbllProductionPlanD LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblPlanBranches LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData TblBranchesData_1 ON dbo.TblPlanBranches.BranchID = TblBranchesData_1.branch_id ON"
MySQL = MySQL & "                       dbo.TbllProductionPlan.TbllProductionPlanD = dbo.TblPlanBranches.planid LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData TblBranchesData_3 ON dbo.TbllProductionPlan.BranchId = TblBranchesData_3.branch_id"

MySQL = MySQL & "   Where (dbo.TbllProductionPlan.TbllProductionPlanD =" & val(TxtTbllProductionPlanD.Text) & ")"

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\RepPaln.rpt"
        Else
           StrFileName = App.path & "\REPORTS\REPORTS NEW\RepPalnE.rpt"
        End If

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?CE??I E?C?CE ?????"
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
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

     
    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " EIC?E ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
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
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(12).AddCurrentValueval (lbTotalMente.Caption)
      'If val(Txt_order_no) <> 0 Then
      
      '  RetrivePoNo Txt_order_no, PONo, oorderdate, CBoBasedON
      
    'End If
  ' If CBoBasedON = 1 Then
    ' xReport.ParameterFields(9).AddCurrentValue oorderdate
    ' xReport.ParameterFields(10).AddCurrentValue PONo
    ' Else
'     xReport.ParameterFields(11).AddCurrentValue oorderdate
'     xReport.ParameterFields(12).AddCurrentValue PONo
    ' End If
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
Function check_previous_dev(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from notes where salary=" & year & Month
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev = False
    Else
        check_previous_dev = True
    End If
 
End Function

Function check_previous_dev1(year As String, Month As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then
        check_previous_dev1 = False
    Else
        check_previous_dev1 = True
    End If
 
End Function

Function Create_dev()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim x As Integer
    Dim rs As ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
         MsgBox "Branch Not Created", vbCritical
        End If
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Else
             MsgBox "Salaries Account Not Selected For this branch", vbCritical
            End If
            GoTo ErrTrap
         
        End If
    End If
        
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "

    Dim StrSQL As String
    Set rs = New ADODB.Recordset
    StrSQL = "select * From Notes where NoteType=66 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = Null
    rs("Remark").value = Msg

    rs("NoteType").value = 66
    rs("NoteDate").value = Date
    rs("UserID").value = user_id
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With
 
 If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    Else
    MsgBox "Entry Created", vbInformation
   End If
    
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
    Else
    MsgBox "An Error Occur while saving data ", vbExclamation
    End If
  
End Function

Function Create_dev1()
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim x As Integer
    Dim rs As ADODB.Recordset
        
    Account_Code_dynamic = get_account_code_branch(16, my_branch)

    If Account_Code_dynamic = "NO branch" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        Else
         MsgBox "Branch Not Created ", vbCritical
        End If
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
    If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Else
            MsgBox "Salar", vbCritical
    End If
            
            GoTo ErrTrap
         
        End If
    End If
        
    'StrAccountCode = Account_Code_dynamic
        
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .Rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With

    Set rs = New ADODB.Recordset
    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
 
    rs("voucher_id").value = LngDevID
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

    dcproject.Text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CmdOk_Click
End Sub

Private Sub ALLButton3_Click()
 
End Sub

Private Sub CboPayMentType_Change()
 
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .Rows = 2
            .Clear flexClearScrollable
        End With

    End If

End Sub

Private Sub CmbMonth_Click()
    CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()

End Sub

Function create_report_data()

End Function

Private Sub CmdPrint_Click()
    On Error Resume Next
    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
End Sub

Private Sub Del_Trans()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If TxtTbllProductionPlanD.Text <> "" Then
    
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtTbllProductionPlanD.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
Else
        Msg = "Will delete process no. " & CHR(13)
        Msg = Msg + (TxtTbllProductionPlanD.Text) & CHR(13)
        Msg = Msg + " are you sure you want to delete"

End If

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Cn.Execute "delete TbllProductionPlanDetails where TbllProductionPlanD=" & val(Me.TxtTbllProductionPlanD.Text)
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    '   XPTxtCurrent.Caption = 0
                    '   XPTxtCount.Caption = 0
                    
                         Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            
                 Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            
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
        Msg = "Process Not Avilable...there is no data"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        Else
           Msg = "Cant delete this record for data Integration" & CHR(13) & "there is data not connected "
    End If
        
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

  On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
        If Trim(Me.Dcbranch.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·›—⁄..!!"
            Else
             Msg = "Select Branch Firstly"
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Dcbranch.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
    End If

If ListGroupSelected.ListCount = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ ›—Ê⁄/›—Ê⁄ «Ê·« ", vbCritical
            Else
                  MsgBox "Select Branch First ", vbCritical
            End If
Exit Sub
End If

Dim i As Integer
    With Me.Grid

        For i = 1 To .Rows - 1

            If val(.TextMatrix(i, .ColIndex("Branchid"))) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Õœœ ›—Ê⁄/›—Ê⁄ «Ê·« ", vbCritical
            Else
                  MsgBox "Select Branch First ", vbCritical
            End If
               
            End If
            
            '
        Next i

    End With
    

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
                Me.TxtTbllProductionPlanD.Text = CStr(new_id("TbllProductionPlan", "TbllProductionPlanD", "", True))
       
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TbllProductionPlanDetails where TbllProductionPlanD=" & val(Me.TxtTbllProductionPlanD.Text)
   
   StrSQL = "Delete From TblPlanBranches Where planId=" & val(Me.TxtTbllProductionPlanD.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
               
               
    End If
    rs("BranchId").value = IIf(Me.Dcbranch.Text = "", Null, Me.Dcbranch.BoundText)
    rs("TbllProductionPlanD").value = TxtTbllProductionPlanD.Text
    '
    rs("CustomerId").value = IIf(Me.DcbCustomer.BoundText = "", Null, Me.DcbCustomer.BoundText)
    rs("FromDate").value = dbFromDate.value
    rs("Todate").value = dbTodate.value
     rs("DbFrom").value = DbFrom.value
    rs("DBTo").value = DBTo.value
    
      rs("planType").value = val(CplanType.ListIndex)
      
    rs("Remarks").value = IIf(Me.TxtRemarks.Text = "", "", Me.TxtRemarks.Text)
  rs("OldPlanNo").value = IIf(Me.TxtOldPlanNo.Text = "", "", Me.TxtOldPlanNo.Text)
  
    If ChkLocked.value = vbChecked Then
        rs("Locked").value = 1
    Else
        rs("Locked").value = 0
    End If
If opt(0).value = True Then
 rs("Opt").value = 0
ElseIf opt(1).value = True Then
 rs("Opt").value = 1
ElseIf opt(2).value = True Then
 rs("Opt").value = 2
ElseIf opt(3).value = True Then
 rs("Opt").value = 3
End If
''// 05 06 2015
 If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
    Else
        rs("CashCustomerName").value = Null
    End If
 rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
 rs("BaseSanad").value = IIf(CBoBaseSanad.ListIndex = -1, Null, val(CBoBaseSanad.ListIndex))
 rs("NoteSerial").value = Trim$(Me.TxtNoteSerial.Text)
    rs.update
 
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TbllProductionPlanDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
 
    With Me.Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
         
                RsDev.AddNew
                RsDev("TbllProductionPlanD").value = Me.TxtTbllProductionPlanD.Text
            
                RsDev("ItemId").value = val(.TextMatrix(i, .ColIndex("ItemId")))
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                
               
                RsDev("BranchId").value = val(.TextMatrix(i, .ColIndex("BranchId")))
                RsDev("Carid").value = val(.TextMatrix(i, .ColIndex("Carid")))
                RsDev("Driverid").value = val(.TextMatrix(i, .ColIndex("Driverid")))
                 RsDev("Remarks").value = .TextMatrix(i, .ColIndex("remarks"))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                 RsDev("Cunt").value = val(.TextMatrix(i, .ColIndex("Cunt")))
                RsDev("Discount").value = val(.TextMatrix(i, .ColIndex("Discount")))
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
 
    RsDev.Close
    'save Groups
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TbllProductionPlanDetailsGroups", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    With Me.Grid1

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("GroupID")) <> "" Then
         
                RsDev.AddNew
                RsDev("TbllProductionPlanD").value = Me.TxtTbllProductionPlanD.Text
            
                RsDev("GroupID").value = val(.TextMatrix(i, .ColIndex("GroupID")))
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                 RsDev("Cunt").value = val(.TextMatrix(i, .ColIndex("Cunt")))
                RsDev("Discount").value = val(.TextMatrix(i, .ColIndex("Discount")))
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
    
    
    Dim RsEmployee As New ADODB.Recordset
Set RsEmployee = Nothing
         
      
      
      
            If ListGroupSelected.ListCount <> 0 Then
             
                        RsEmployee.Open "TblPlanBranches", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                        
             
            
                            For i = 0 To ListGroupSelected.ListCount - 1
                            RsEmployee.AddNew
                                 RsEmployee("BranchID").value = ListGroupSelected.ItemData(i)
                                    RsEmployee("PlanId").value = val(TxtTbllProductionPlanD.Text)
                              
                            RsEmployee.update
                        Next i
            
  RsEmployee.Close
            End If
            

  
  
    Cn.CommitTrans
    BeginTrans = False
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"
        
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
        Else
            Msg = " Process Data Saved Successfully  " & CHR(13)
            Msg = Msg + "do you want to add another data"
        End If
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
            MsgBox "Updates Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
            
            '  Fg_Journal.Enabled = False
    End Select

    TxtModFlg.Text = "R"
    'End If

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
        Msg = "Process Data Can't be Saved " & CHR(13)
        Msg = Msg + "Invalid Data Entered " & CHR(13)
        Msg = Msg + "Please Be Sure from Data accuracy"
     
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    
 If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
      Msg = "Sorry.... An error occur while saving data " & CHR(13)
    End If
    
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Cmd_Click(Index As Integer)
     On Error GoTo ErrTrap

    Select Case Index
  Case 10

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report
        Case 0

 
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me

            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            Grid.Enabled = True
            Option2.value = True
            Option4.value = True
            Me.Dcbranch.BoundText = Current_branch
CplanType.ListIndex = val(Me.lblPlantype.Caption)
DbFrom.value = Date
DBTo.value = Date
  ListGroupSelected.Clear
opt(0).value = True

Me.CplanType.ListIndex = lblPlantype.Caption

        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
            CuurentLogdata

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
Load PlanSearch
PlanSearch.show
           ' Load FrmNotesSearch
           ' FrmNotesSearch.SearchType = 3
           ' FrmNotesSearch.Show vbModal

        Case 6
            Unload Me

        Case 7
            '   ViewDataList
            addrowGroups
    
        Case 8
            RemoveGridRowGroup
    
        Case 20
        If ListGroupSelected.ListCount = 0 Then
                  If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "Õœœ ›—Ê⁄/›—Ê⁄ «Ê·« ", vbCritical
                    Else
                          MsgBox "Select Branch First ", vbCritical
                    End If
        Exit Sub
        End If


            addrow

Case 9
If ListGroupSelected.ListCount = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "Õœœ ›—Ê⁄/›—Ê⁄ «Ê·« ", vbCritical
                Else
                      MsgBox "Select Branch First ", vbCritical
                End If
Exit Sub
End If

RetriveOrder dbFromDate.value, DBTo.value

        Case 21
            RemoveGridRow
            
            Case 11
            Me.Grid.Clear flexClearScrollable, flexClearEverything
 Grid.Rows = 1
 
    End Select

    Exit Sub
ErrTrap:

End Sub
Public Sub RetriveOrderTrans(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDev As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
    
    Grid.Refresh
DcbCustomer.Text = ""
TxtCashCustomerName.Text = ""
DCboStoreName.Text = ""

StrSQL = " SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.*"
StrSQL = StrSQL & " FROM         dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " where  dbo.Transactions.Transaction_Type=" & Transaction_Type & "  and dbo.Transactions.NoteSerial1='" & order_no & "'"



    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DcbCustomer.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
       
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
'        Me.Dcbranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)
      If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If
    
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

StrSQL = StrSQL + "order by dbo.Transaction_Details.id"

    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
With Grid
    If Not (RsDev.EOF Or RsDev.BOF) Then
        .Rows = RsDev.RecordCount + 1

        For i = 1 To RsDev.RecordCount
                        .TextMatrix(i, .ColIndex("Ser")) = i
                        .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(RsDev("ItemId").value), "", RsDev("ItemId").value)
                        .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
                        .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), "", RsDev("UnitId").value)
                        .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
                        .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)
                        .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("showqty").value), 0, RsDev("showqty").value)
                        .TextMatrix(i, .ColIndex("Cunt")) = IIf(IsNull(RsDev("showPrice").value), 0, RsDev("showPrice").value)
                        .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("Cunt"))) * val(.TextMatrix(i, .ColIndex("Price")))
                      
      ' .TextMatrix(i, .ColIndex("Driverid")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
      '         .TextMatrix(i, .ColIndex("Carid")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
      '          .TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(RsDev("BoardNO").value), "", RsDev("BoardNO").value)
                   
If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                .TextMatrix(i, .ColIndex("branchname")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                ' .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
Else
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitNamee").value), "", RsDev("UnitNamee").value)
                .TextMatrix(i, .ColIndex("branchname")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                ' .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemNamee").value), "", RsDev("ItemNamee").value)

End If

  
               
                 
               
            
              '  .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            ''///
       
            Debug.Print i

            If Grid.Rows > 10 Then
                If i = 8 Then Grid.Refresh
            End If
RsDev.MoveNext
        Next i

    End If

End With


    Screen.MousePointer = vbDefault
    
'    XPTxtCurrent.Caption = rs.AbsolutePosition
'    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Public Sub RetriveOrder(FromDate As Date, ToDate As Date)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    Dim Transaction_Type As Integer
Dim branchStr As String
    On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 2
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Refresh
    
If opt(0).value = True Then
Transaction_Type = 6
ElseIf opt(1).value = True Then
Transaction_Type = 21
ElseIf opt(2).value = True Then
Transaction_Type = 54
ElseIf opt(3).value = True Then

ElseIf opt(4).value = True Then
Transaction_Type = 38

End If
Dim i As Integer
branchStr = branchStr
For i = 0 To ListGroupSelected.ListCount - 1
        If i = ListGroupSelected.ListCount - 1 Then
                branchStr = branchStr & ListGroupSelected.ItemData(i)
        Else
                branchStr = branchStr & ListGroupSelected.ItemData(i) & ","
        End If

Next i

StrSQL = " SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.Transactions.Approved, dbo.Transaction_Details.Item_ID, "
  StrSQL = StrSQL & " dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ShowQty, dbo.Transactions.BranchId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
  StrSQL = StrSQL & " dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.Transaction_Date, dbo.TblBranchesData.branch_name,"
  StrSQL = StrSQL & " dbo.TblBranchesData.branch_nameE"
  StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
  StrSQL = StrSQL & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
  StrSQL = StrSQL & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
  StrSQL = StrSQL & " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID INNER JOIN"
  StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
 
       StrSQL = StrSQL & " where  Transaction_Type=" & Transaction_Type
    If Not IsNull(FromDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date >=" & SQLDate(CDate(FromDate), True) & ""
    End If

    If Not IsNull(ToDate) Then
        StrSQL = StrSQL + " And dbo.Transactions.Transaction_Date <=" & SQLDate(CDate(ToDate), True) & ""
    End If
 
 If branchStr <> "" Then
    StrSQL = StrSQL + " And dbo.Transactions.BranchId IN(" & branchStr & ")"
 End If
 
 

    Set RsDetails = New ADODB.Recordset
   RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
    If RsDetails.RecordCount < 1 Then
 
        Exit Sub
    Else
   
    End If

    If RsDetails.EOF Or RsDetails.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
 
   

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
       Grid.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
         
           Grid.TextMatrix(Num, Grid.ColIndex("ItemID")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
           Grid.TextMatrix(Num, Grid.ColIndex("ItemCode")) = IIf(IsNull(RsDetails("ItemCode")), "", (RsDetails("ItemCode").value))
          If SystemOptions.UserInterface = ArabicInterface Then
           Grid.TextMatrix(Num, Grid.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemName")), "", (RsDetails("ItemName").value))
            Grid.TextMatrix(Num, Grid.ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
     
         Else
         Grid.TextMatrix(Num, Grid.ColIndex("itemname")) = IIf(IsNull(RsDetails("ItemNamee")), "", (RsDetails("ItemNamee").value))
        Grid.TextMatrix(Num, Grid.ColIndex("UnitName")) = IIf(IsNull(RsDetails("UnitNamee")), "", (RsDetails("UnitNamee").value))
         
         End If
  
           Grid.TextMatrix(Num, Grid.ColIndex("unitid")) = IIf(IsNull(RsDetails("UnitId")), "", (RsDetails("UnitId").value))
            Grid.TextMatrix(Num, Grid.ColIndex("Branchid")) = IIf(IsNull(RsDetails("BranchId")), "", (RsDetails("BranchId").value))
           If SystemOptions.UserInterface = ArabicInterface Then
           Grid.TextMatrix(Num, Grid.ColIndex("Branchname")) = IIf(IsNull(RsDetails("branch_name")), "", (RsDetails("branch_name").value))
          Else
         Grid.TextMatrix(Num, Grid.ColIndex("Branchname")) = IIf(IsNull(RsDetails("branch_nameE")), "", (RsDetails("branch_nameE").value))
          End If
        
        'ShowQty
                   Grid.TextMatrix(Num, Grid.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))


        
            RsDetails.MoveNext
   
   

            If Grid.Rows > 10 Then
                If Num = 8 Then Grid.Refresh
            End If

        Next Num

    End If

   
    Screen.MousePointer = vbDefault
 
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub RemoveGridRowGroup()

    With Me.Grid1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Function addrow()

    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Integer
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer

    If Option2.value = True Then
        If dcitems.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— «·’‰›  ...!!!"
            Else
                Msg = "must Specify item Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If

        wherestr = "  where ItemID= " & val(dcitems.BoundText)
    End If

    sql = "Select * from TblItems "

    If wherestr <> "" Then
        sql = sql & wherestr
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    With Grid
 
        lastrow = .Rows
    
        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + lastrow
            Rs3.MoveFirst
         
            For i = lastrow To Rs3.RecordCount + lastrow - 1
           
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                LngItemID = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                       
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs3.Fields("ItemCode").value), "", Rs3.Fields("ItemCode").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3.Fields("ItemName").value), "", Rs3.Fields("ItemName").value)
                Else
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3.Fields("ItemNamee").value), "", Rs3.Fields("ItemName").value)
                End If
                       
                'lllllllllllllll
                StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName , TblUnites.UnitNamee  "
                StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
                StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                 
                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsUnit.RecordCount > 0 Then
                    RsUnit.MoveFirst
                    .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
                    Else
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitNamee").value), "", RsUnit.Fields("UnitNamee").value)
                    End If
               
                End If
 ' If i = 1 Then
                'If ListGroupSelected.ListIndex > -1 Then
                          .TextMatrix(i, .ColIndex("Branchname")) = ListGroupSelected.List(0)
                            .TextMatrix(i, .ColIndex("Branchid")) = ListGroupSelected.ItemData(0)
                       '   End If
        '    End If
                RsUnit.Close
                       
                Rs3.MoveNext
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

    ReLineGrid

End Function

Function addrowGroups()

    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Integer
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer

    If Option4.value = True Then
        If DCGroup.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— „Ã„Ê⁄Â  ...!!!"
            Else
                Msg = "must Specify item Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If

        wherestr = "  where GroupID= " & val(DCGroup.BoundText)
    End If

    sql = "Select * from Groups "

    If wherestr <> "" Then
        sql = sql & wherestr
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    With Grid1
 
        lastrow = .Rows
    
        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + lastrow
            Rs3.MoveFirst
         
            For i = lastrow To Rs3.RecordCount + lastrow - 1
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs3.Fields("GroupID").value), "", Rs3.Fields("GroupID").value)
                LngItemID = IIf(IsNull(Rs3.Fields("GroupID").value), "", Rs3.Fields("GroupID").value)
                       
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs3.Fields("Fullcode").value), "", Rs3.Fields("Fullcode").value)
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs3.Fields("GroupName").value), "", Rs3.Fields("GroupName").value)
                       
                'lllllllllllllll
                '     StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                '   StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & _
                '   "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                '   StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
                '   StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                 
                StrSQL = "SELECT TblUnites.UnitID, TblUnites.UnitName "
                StrSQL = StrSQL + " FROM TblUnites  "
                
                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsUnit.RecordCount > 0 Then
                    RsUnit.MoveFirst
                    .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
               
                End If

                RsUnit.Close
                       
                Rs3.MoveNext
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

    ReLineGrid

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Dcdep_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub

Function SHow_grig_col()
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

Private Sub CmdRemove_Click()
    Dim x As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        x = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
    If SystemOptions.UserInterface = ArabicInterface Then
        x = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
        Else
          x = MsgBox("are you sure you want to delete", vbCritical + vbYesNo)
        End If
    End If

    If x = vbNo Then Exit Sub
    
    If Grid.Rows > 1 Then
        If Grid.Rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.Rows > 1 Then
                If Me.Grid.Row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.Row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub

 

Private Sub DataCombo5_Click(Area As Integer)

End Sub

'Private Sub DataCombo5_KeyUp(KeyCode As Integer, Shift As Integer)
Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
       FrmCustemerSearch.searchtype = 17
        FrmCustemerSearch.show vbModal
    End If
    
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        'Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
 
    End If
        
End Sub
 

Private Sub DcbCustomer_Change()
    
    Dim Msg As String
    Dim RsTemp  As ADODB.Recordset
    Dim StrSQL As String

    Dim Fullcode As String
    GetCustomersDetail val(DcbCustomer.BoundText), , Fullcode, 1
    TxtSearchCode.Text = Fullcode

End Sub

Private Sub DCboStoreName_Change()
TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))


End Sub

Private Sub dcitems_Change()
    Me.TxtItemCode.Text = GetItemCode(val(Me.dcitems.BoundText))

End Sub

Private Sub dcitems_Click(Area As Integer)
dcitems_Change
End Sub

Private Sub dcproject_Click(Area As Integer)

    If dcproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub Dcterm_Click(Area As Integer)

    If Dcterm.BoundText = "" Then Exit Sub

    My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
    fill_combo dcopr, My_SQL
End Sub

Private Sub Label5_Click()

    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If

End Sub
Function FillMylist()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
  

    sql = " SELECT * from  TblBranchesData "
 
 If SystemOptions.UserInterface = ArabicInterface Then
sql = sql & " order by  branch_name"
Else
sql = sql & " order by  branch_name"
End If
 
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
'    ListGroupSelected.Clear

    If rs.RecordCount > 0 Then

        For i = 1 To rs.RecordCount
             
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("branch_id").value
            rs.MoveNext
        Next i

    End If

    rs.Close

End Function

Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub

Private Sub Label7_Click()
    Dim i As Integer
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i

End Sub

Private Sub Label8_Click()
If ListGroupAll.ListIndex = -1 Then Exit Sub
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
             
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
            
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With

    With Me.Grid1
        Set .WallPaper = GrdBack.Picture
     
    End With
    If SystemOptions.UserInterface = ArabicInterface Then
 With Me.CBoBaseSanad
        .Clear
        .AddItem "›« Ê—… „»Ì⁄« "
          .AddItem "√„— »Ì⁄"

        .AddItem " ÿ·» œ«Œ·Ì"
    End With
    Else
    With Me.CBoBaseSanad
        .Clear
        .AddItem "Sales Invoices"
          .AddItem "Sell Order"

        .AddItem " Internal Order"
    End With
    End If
    
With CplanType
If SystemOptions.UserInterface = EnglishInterface Then

.Clear
.AddItem (" Purchase Plan ")
.AddItem (" Production Plan ")
.AddItem (" Sales Plan ")
.AddItem ("Shipping Plan")
Else
.Clear
.AddItem ("Œÿ… „‘ —Ì« ")
.AddItem (" ŒÿÂ «‰ «Ã ")
.AddItem (" Œÿ… „»Ì⁄«  ")
.AddItem ("Œÿ… ‘Õ‰")
End If
End With

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

 '   Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    Dcombos.GetItemsNames dcitems
    Dcombos.GetItemSGroups DCGroup
    Dcombos.GetBranches Me.Dcbranch
       Dcombos.GetCustomersSuppliers 1, Me.DcbCustomer
    Dcombos.GetStores Me.DCboStoreName
    FillMylist
    
    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TbllProductionPlan  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
   opt(4).Caption = "Inner Request"
    
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"
opt(3).RightToLeft = True
opt(3).Caption = "Previous plan"

opt(2).RightToLeft = True
opt(2).Caption = "Shipping Requests"
opt(1).RightToLeft = True
opt(1).Caption = "Sales Invoices "
opt(0).RightToLeft = True
Frame3.Caption = ""
opt(0).Caption = "Sell Order "

    Me.Caption = "Plan"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "Plan No"
    lbl(11).Caption = "From"
    Label12.Caption = "Selected Branch"
    lbl(12).Caption = "To"
    FrameBranch.Caption = "Select Branch"
    Frame2.Caption = "Based On"
    lbl(13).Caption = "Plan Type"
    lbl(5).Caption = "Start "
    lbl(2).Caption = "End "
    lbl(0).Caption = "Branch"
    lbl(3).Caption = "Remarks"
    Cmd(9).Caption = "Show Data"
    ChkLocked.Caption = "Locked"
    Cmd(7).Caption = "Add"
    Cmd(8).Caption = "Remove"
    Option1.Caption = "All Item"
    Option2.Caption = "Select Item"
    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Remove"
Label2.Caption = "Plan No."
    CmdRemove.Caption = "Remove Line"
 
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "ser"
        .TextMatrix(0, .ColIndex("UnitName")) = "UnitName"
        .TextMatrix(0, .ColIndex("Price")) = "Qty"
        .TextMatrix(0, .ColIndex("Cunt")) = "Price"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ItemCode"
        .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
        .TextMatrix(0, .ColIndex("Branchname")) = "BranchName"
        .TextMatrix(0, .ColIndex("Carname")) = "CarName"
        .TextMatrix(0, .ColIndex("DriverName")) = "DriverName"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("Total")) = "Total"
    End With

   ' With Me.Grid
   '     .TextMatrix(0, .ColIndex("ser")) = "I"
   '     .TextMatrix(0, .ColIndex("ItemCode")) = "ItemCode"
   '     .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
   '     .TextMatrix(0, .ColIndex("UnitName")) = "UnitName"
   '     .TextMatrix(0, .ColIndex("Price")) = "Price"
   '     .TextMatrix(0, .ColIndex("discount")) = "Discount"
   ' End With

    Me.C1Tab1.TabCaption(1) = "Groups"
    Me.C1Tab1.TabCaption(0) = "Items"
   Cmd(11).Caption = "Delete All"
   lbl(16).Caption = "No."
   Cmd(10).Caption = "Print"
    Label28.Caption = "Cash Clien "
    lbl(15).Caption = "Spec. Store "
    lbl(17).Caption = "Spec. Bill"
    lbl(42).Caption = "Spec. Client"
    
End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .Rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Ser")) = i
                ',DepartmentID,project_id
            
                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
            
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
           
                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                rs.MoveNext
            
            Next

            rs.Close
        End If

        .Rows = .Rows + 1
        If SystemOptions.UserInterface = ArabicInterface Then
        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        Else
           .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "Total"
        End If
        
        .IsSubtotal(.Rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

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
     
                SaveData

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim code  As String

    With Grid

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
                .TextMatrix(Row, .ColIndex("UnitID")) = code
  
   Case "Branchname"
                 code = .ComboData
                .TextMatrix(Row, .ColIndex("BranchId")) = code
                    
   Case "Carname"
                code = .ComboData
               .TextMatrix(Row, .ColIndex("Carid")) = code
                 
                
   Case "DriverName"
                 code = .ComboData
                .TextMatrix(Row, .ColIndex("Driverid")) = code
                
                
        End Select
   
        If Row = .Rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With
 
    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    'Grid.TextMatrix(Row, Grid.ColIndex("Code"))
    'Grid.TextMatrix(Row, Grid.ColIndex("Name"))
    If Col = Grid.ColIndex("ItemCode") Or Col = Grid.ColIndex("ItemName") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , , , Me.TxtTbllProductionPlanD
    ElseIf Col = Grid.ColIndex("UnitName") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), Grid.TextMatrix(Row, Grid.ColIndex("UnitName")), , , , , , , , , , Me.TxtTbllProductionPlanD
    ElseIf Col = Grid.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , (Grid.TextMatrix(Row, Grid.ColIndex("Price"))), , , , , , , , Me.TxtTbllProductionPlanD
    ElseIf Col = Grid.ColIndex("Discount") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , Grid.TextMatrix(Row, Grid.ColIndex("Discount")), , Me.TxtTbllProductionPlanD

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("Cunt"))) * val(.TextMatrix(i, .ColIndex("Price")))
  
            End If

        Next i
   
    End With

    With Me.Grid1

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("GroupID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If

   If .ColKey(Col) <> "Branchname" Then
       
            .ComboList = ""
        End If
           If .ColKey(Col) <> "Carname" Then
       
            .ComboList = ""
        End If
   If .ColKey(Col) <> "DriverName" Then
       
            .ComboList = ""
        End If
        
    End With

End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
'    Dim StrAccountType As String
    Dim StrComboList As String
  '  Dim Msg As String
    Dim LngItemID As Integer
     Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "UnitName"

                LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemId")))

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName ,TblUnites.UnitNamee "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        Else
                         MyStrList = .BuildComboList(rs, "UnitNamee", "UnitID")
                        End If
                        
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
           Case "Branchname"
     If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "  select branch_id,branch_name from TblBranchesData   "
                Else
                    StrSQL = "  select branch_id,branch_namee from TblBranchesData   "
           End If
         
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = .BuildComboList(rs, "branch_name", "branch_id")
                Else
                    StrComboList = .BuildComboList(rs, "branch_namee", "branch_id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList



      Case "Carname"
                    StrSQL = "  select id,BoardNO from TblCarsData"
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = .BuildComboList(rs, "BoardNO", "id")
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList



         
      Case "DriverName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
     
 
         If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = "SELECT     dbo.tblCarDrivers.EmpID, dbo.TblEmployee.Emp_Name"
            StrSQL = StrSQL & " FROM         dbo.tblCarDrivers LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.tblCarDrivers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.tblCarDrivers.EmpID = dbo.TblEmployee.Emp_ID  "
  StrSQL = StrSQL + " Order By Emp_Name ASC"
    Else
            StrSQL = "SELECT     dbo.tblCarDrivers.EmpID, dbo.TblEmployee.Emp_Namee"
            StrSQL = StrSQL & " FROM         dbo.tblCarDrivers LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.tblCarDrivers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
            StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.tblCarDrivers.EmpID = dbo.TblEmployee.Emp_ID  "
            StrSQL = StrSQL + " Order By Emp_Namee  ASC"

    End If
    
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                        StrComboList = .BuildComboList(rs, "Emp_Name", "EmpID")
                Else
                        StrComboList = .BuildComboList(rs, "Emp_Namee", "EmpID")
                End If
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
        End Select
    End With
End Sub

Public Sub reterivePlan(Optional Lngid As Long = 0)
 
   Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
    
StrSQL = "   SELECT     dbo.TbllProductionPlanDetails.TbllProductionPlanD, dbo.TbllProductionPlanDetails.UnitID, dbo.TbllProductionPlanDetails.ItemID,"
StrSQL = StrSQL & "    dbo.TbllProductionPlanDetails.Discount, dbo.TbllProductionPlanDetails.Price, dbo.TblUnites.UnitName, dbo.TblItems.ItemName, dbo.TblItems.ItemCode,"
StrSQL = StrSQL & "    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "    dbo.TblEmployee.Emp_Namee , dbo.TblCarsData.BoardNO"
StrSQL = StrSQL & "   , dbo.TblCarsData.id, dbo.TblBranchesData.branch_id, dbo.TblEmployee.Emp_ID  FROM         dbo.TbllProductionPlanDetails INNER JOIN"
StrSQL = StrSQL & "    dbo.TblItems ON dbo.TbllProductionPlanDetails.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblBranchesData ON dbo.TbllProductionPlanDetails.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblEmployee ON dbo.TbllProductionPlanDetails.Driverid = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblCarsData ON dbo.TbllProductionPlanDetails.Carid = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblUnites ON dbo.TbllProductionPlanDetails.UnitID = dbo.TblUnites.UnitID"
StrSQL = StrSQL & "  where TbllProductionPlanD=" & Lngid
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(RsDev("ItemId").value), "", RsDev("ItemId").value)
            
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
            
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), "", RsDev("UnitId").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
                      .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(RsDev("branch_id").value), "", RsDev("branch_id").value)
       .TextMatrix(i, .ColIndex("Driverid")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
               .TextMatrix(i, .ColIndex("Carid")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
                .TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(RsDev("BoardNO").value), "", RsDev("BoardNO").value)
                   
If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branchname")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
                 .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
Else
                .TextMatrix(i, .ColIndex("branchname")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
                 .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)


End If

  
               
                   
                   
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
            
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    End If

    RsDev.Close

End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

     On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
    Grid1.Clear flexClearScrollable, flexClearEverything
    Grid1.Rows = 1

    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
 
            
       If Lngid <> 0 Then
            rs.find "TbllProductionPlanD=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
 
    Me.TxtTbllProductionPlanD.Text = IIf(IsNull(rs("TbllProductionPlanD").value), "", rs("TbllProductionPlanD").value)
     Me.Dcbranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    dbFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dbTodate.value = IIf(IsNull(rs("Todate").value), Date, rs("Todate").value)
    DbFrom.value = IIf(IsNull(rs("DbFrom").value), Date, rs("DbFrom").value)
    DBTo.value = IIf(IsNull(rs("DBTo").value), Date, rs("DBTo").value)
    TxtOldPlanNo.Text = IIf(IsNull(rs("OldPlanNo").value), "", rs("OldPlanNo").value)
'  rs("OldPlanNo").value = IIf(Me.TxtOldPlanNo.text = "", "", Me.TxtOldPlanNo.text)

    DcbCustomer.BoundText = IIf(IsNull(rs("CustomerId").value), "", rs("CustomerId").value)
    CplanType.ListIndex = IIf(IsNull(rs("planType").value), 0, rs("planType").value)

    TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
 'rs("Opt").value = 1
    If IsNull(rs("Locked").value) Then
        ChkLocked.value = vbUnchecked
    Else
        ChkLocked.value = rs("Locked").value
    End If
If rs("Opt").value = 0 Then
Me.opt(0).value = True
ElseIf rs("Opt").value = 1 Then
Me.opt(1).value = True
  ElseIf rs("Opt").value = 2 Then
Me.opt(2).value = True
ElseIf rs("Opt").value = 3 Then
Me.opt(3).value = True
End If '''

''// 05 06 2015
If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If
Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
Me.CBoBaseSanad.ListIndex = IIf(IsNull(rs("BaseSanad").value), -1, rs("BaseSanad").value)
TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
StrSQL = " SELECT     dbo.TbllProductionPlanDetails.TbllProductionPlanD, dbo.TbllProductionPlanDetails.UnitID, dbo.TbllProductionPlanDetails.ItemID,"
   StrSQL = StrSQL & "                    dbo.TbllProductionPlanDetails.Discount, dbo.TbllProductionPlanDetails.Price, dbo.TblUnites.UnitName, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.ItemCode,"
   StrSQL = StrSQL & "                    dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
 StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TblCarsData.BoardNO, dbo.TblCarsData.id, dbo.TblBranchesData.branch_id, dbo.TblEmployee.Emp_ID,"
 StrSQL = StrSQL & "                      dbo.TbllProductionPlanDetails.Remarks , dbo.TbllProductionPlanDetails.Cunt"
StrSQL = StrSQL & "  FROM         dbo.TbllProductionPlanDetails INNER JOIN"
 StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TbllProductionPlanDetails.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TbllProductionPlanDetails.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TbllProductionPlanDetails.Driverid = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCarsData ON dbo.TbllProductionPlanDetails.Carid = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.TbllProductionPlanDetails.UnitID = dbo.TblUnites.UnitID"
StrSQL = StrSQL & " Where (dbo.TbllProductionPlanDetails.TbllProductionPlanD =" & val(TxtTbllProductionPlanD.Text) & ")"

    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(RsDev("ItemId").value), "", RsDev("ItemId").value)
            
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
             .TextMatrix(i, .ColIndex("remarks")) = IIf(IsNull(RsDev("Remarks").value), "", RsDev("Remarks").value)
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                Else
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemNameE").value), "", RsDev("ItemNamee").value)
                End If
                
                .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), "", RsDev("UnitId").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
                      .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(RsDev("branch_id").value), "", RsDev("branch_id").value)
       .TextMatrix(i, .ColIndex("Driverid")) = IIf(IsNull(RsDev("Emp_ID").value), "", RsDev("Emp_ID").value)
               .TextMatrix(i, .ColIndex("Carid")) = IIf(IsNull(RsDev("id").value), "", RsDev("id").value)
                .TextMatrix(i, .ColIndex("CarName")) = IIf(IsNull(RsDev("BoardNO").value), "", RsDev("BoardNO").value)
                   
If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("branchname")) = IIf(IsNull(RsDev("branch_name").value), "", RsDev("branch_name").value)
                 .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(RsDev("Emp_Name").value), "", RsDev("Emp_Name").value)
Else
                .TextMatrix(i, .ColIndex("branchname")) = IIf(IsNull(RsDev("branch_namee").value), "", RsDev("branch_namee").value)
                 .TextMatrix(i, .ColIndex("DriverName")) = IIf(IsNull(RsDev("Emp_Namee").value), "", RsDev("Emp_Namee").value)


End If

  
               
                   .TextMatrix(i, .ColIndex("Cunt")) = IIf(IsNull(RsDev("Cunt").value), 0, RsDev("Cunt").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
                .TextMatrix(i, .ColIndex("total")) = val(.TextMatrix(i, .ColIndex("Cunt"))) * val(.TextMatrix(i, .ColIndex("Price")))
            
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    End If

    RsDev.Close
    'fill Group grid
    StrSQL = " SELECT     dbo.TbllProductionPlanDetailsGroups.TbllProductionPlanD, dbo.TbllProductionPlanDetailsGroups.UnitID, dbo.TbllProductionPlanDetailsGroups.GroupID, dbo.TbllProductionPlanDetailsGroups.Discount, "
    StrSQL = StrSQL & "     dbo.TbllProductionPlanDetailsGroups.Price , dbo.TblUnites.unitname, dbo.Groups.GroupName, dbo.Groups.Fullcode"
    StrSQL = StrSQL & " FROM         dbo.TbllProductionPlanDetailsGroups INNER JOIN"
    StrSQL = StrSQL & " dbo.Groups ON dbo.TbllProductionPlanDetailsGroups.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblUnites ON dbo.TbllProductionPlanDetailsGroups.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  where TbllProductionPlanD=" & val(Me.TxtTbllProductionPlanD.Text)
 
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(RsDev("GroupID").value), "", RsDev("GroupID").value)
            
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
            
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(RsDev("GroupName").value), "", RsDev("GroupName").value)
                .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), "", RsDev("UnitId").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
            
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
 
 
    ListGroupSelected.Clear
    

 
    
 
 
 
StrSQL = " SELECT     TOP 100 PERCENT dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
StrSQL = StrSQL & " FROM         dbo.TblPlanBranches INNER JOIN"
StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblPlanBranches.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " Where (dbo.TblPlanBranches.planid = " & val(Me.TxtTbllProductionPlanD.Text) & ")"
StrSQL = StrSQL & " ORDER BY dbo.TblPlanBranches.id"


 
Dim RsEmployee As New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
 
                            For i = 0 To RsEmployee.RecordCount - 1
                                
                                
                        
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                       ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_name").value), "", RsEmployee("branch_name").value)
                                Else
                                     ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_nameE").value), "", RsEmployee("branch_nameE").value)
                                End If
                                ListGroupSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("branch_id").value), 0, val(RsEmployee("branch_id").value)))
                
                                RsEmployee.MoveNext
                            Next i
RsEmployee.Close
Set RsEmployee = Nothing

    End If


'*********************************************************************************
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Grid1_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid1

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
         
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
   
        If Row = .Rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With Grid1

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If

    End With

End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "UnitName"

                LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemId")))

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    '        StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                    '        StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & _
                    '        "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    '        StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    '        StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    StrSQL = "SELECT TblUnites.UnitID, TblUnites.UnitName "
                    StrSQL = StrSQL + " FROM TblUnites   "
                
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
        End Select

    End With

End Sub

 

 

Private Sub lblPlantype_Click()
'FileCopy  "c:\my file.txt", "c:\1.txt"
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtItemCode.Text = "" Then
            Me.dcitems.BoundText = ""
        Else
            Me.dcitems.BoundText = GetItemID(Trim$(Me.TxtItemCode.Text))
        End If
    End If
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub TxtNoteSerial_Change()
    Dim Transaction_Type As Integer
    If CBoBaseSanad.ListIndex = 0 Then
        Transaction_Type = 21
   ElseIf CBoBaseSanad.ListIndex = 1 Then
        Transaction_Type = 6
    ElseIf CBoBaseSanad.ListIndex = 2 Then
        Transaction_Type = 38
    End If


    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrderTrans Me.TxtNoteSerial.Text, Transaction_Type
    End If
End Sub

Private Sub TxtNoteSerial_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
          
                
       If CBoBaseSanad.ListIndex = 2 Then
     FrmBuySearch.DealingForm = GridTransType.internalorder
     FrmBuySearch.Index = 8
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
            FrmBuySearch.show vbModal
         ElseIf CBoBaseSanad.ListIndex = 0 Then
     FrmBuySearch.DealingForm = GridTransType.InvoiceTransaction
     FrmBuySearch.Index = 9
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „»Ì⁄«    "
            FrmBuySearch.show vbModal
                ElseIf CBoBaseSanad.ListIndex = 1 Then
    ' FrmBuySearch.DealingForm = GridTransType.InvoiceTransaction
     
     Order_no_search.lblSpecificsearch = 6
     Order_no_search.RetrunType = 16
   
            Order_no_search.Caption = "«·»ÕÀ ⁄‰ «„—  »Ì⁄"
            Order_no_search.Label1(2).Caption = Order_no_search.Caption
            Order_no_search.show
       End If
   End If
End Sub

Private Sub TxtOldPlanNo_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
        If opt(3).value = True Then
        
            reterivePlan (val(TxtOldPlanNo))
        End If
End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
 Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DcbCustomer.BoundText = CUSTID
    End If
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

 On Error GoTo ErrTrap

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
