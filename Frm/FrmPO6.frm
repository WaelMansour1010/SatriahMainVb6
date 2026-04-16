VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmPO6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÿ·»«   œ«Œ·Ì…"
   ClientHeight    =   9780
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   14025
   HelpContextID   =   340
   Icon            =   "FrmPO6.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   14025
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9780
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   14025
      _cx             =   24739
      _cy             =   17251
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   390
         Index           =   3
         Left            =   15
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   8535
         Width           =   13935
         _cx             =   24580
         _cy             =   688
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Left            =   14085
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   30
            Width           =   1275
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   4170
            TabIndex        =   12
            Top             =   45
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   6660
            TabIndex        =   127
            Top             =   0
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   270
            Index           =   49
            Left            =   8370
            TabIndex        =   129
            Top             =   60
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   6630
            TabIndex        =   128
            Top             =   30
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label LblDiscountsTotalView 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   9120
            TabIndex        =   124
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label LblDiscountsTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   9135
            TabIndex        =   126
            Top             =   30
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   270
            Index           =   50
            Left            =   10425
            TabIndex        =   125
            Top             =   60
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   270
            Index           =   63
            Left            =   5055
            TabIndex        =   70
            Top             =   120
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   330
            Left            =   15900
            TabIndex        =   69
            Top             =   0
            Visible         =   0   'False
            Width           =   2670
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   360
            Left            =   12090
            TabIndex        =   68
            Top             =   0
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   25
            Left            =   13065
            TabIndex        =   67
            Top             =   60
            Visible         =   0   'False
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã„«·Ì «·ÿ·»"
            Height          =   240
            Index           =   3
            Left            =   14340
            TabIndex        =   18
            Top             =   60
            Width           =   1920
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   0
            Left            =   2925
            TabIndex        =   17
            Top             =   105
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   2
            Left            =   1080
            TabIndex        =   16
            Top             =   105
            Width           =   945
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   240
            Left            =   2235
            TabIndex        =   15
            Top             =   90
            Width           =   720
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   210
            Left            =   90
            TabIndex        =   14
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   285
            Index           =   1
            Left            =   5520
            TabIndex        =   13
            Top             =   60
            Width           =   1050
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2610
         Index           =   0
         Left            =   0
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   525
         Width           =   14010
         _cx             =   24712
         _cy             =   4604
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
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "FrmPO6.frx":038A
            Left            =   7305
            List            =   "FrmPO6.frx":038C
            Style           =   2  'Dropdown List
            TabIndex        =   190
            Top             =   2220
            Width           =   4950
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4815
            TabIndex        =   187
            Top             =   660
            Width           =   825
         End
         Begin VB.ComboBox DcbPeriodsID 
            Height          =   315
            ItemData        =   "FrmPO6.frx":038E
            Left            =   120
            List            =   "FrmPO6.frx":039B
            TabIndex        =   171
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox TxtPeriods 
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
            Height          =   270
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   170
            Top             =   960
            Width           =   825
         End
         Begin VB.ComboBox purchaseType 
            Height          =   315
            ItemData        =   "FrmPO6.frx":03AE
            Left            =   2880
            List            =   "FrmPO6.frx":03B0
            Style           =   2  'Dropdown List
            TabIndex        =   160
            Top             =   960
            Width           =   2730
         End
         Begin VB.TextBox txtempcode 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   11385
            TabIndex        =   158
            Top             =   1920
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   4560
            TabIndex        =   155
            Top             =   1275
            Width           =   1065
         End
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4560
            TabIndex        =   154
            Top             =   1590
            Width           =   1065
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11385
            TabIndex        =   150
            Top             =   1590
            Width           =   825
         End
         Begin VB.TextBox TxtStoreID1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   17985
            TabIndex        =   149
            Top             =   1590
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   2880
            TabIndex        =   146
            Top             =   0
            Width           =   1575
         End
         Begin VB.ComboBox DCOPrType 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   144
            Top             =   0
            Width           =   1815
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmPO6.frx":03B2
            Left            =   4560
            List            =   "FrmPO6.frx":03B4
            Style           =   2  'Dropdown List
            TabIndex        =   142
            Top             =   0
            Width           =   1095
         End
         Begin VB.ComboBox CBOOrderType 
            Height          =   315
            ItemData        =   "FrmPO6.frx":03B6
            Left            =   7305
            List            =   "FrmPO6.frx":03B8
            Style           =   2  'Dropdown List
            TabIndex        =   140
            Top             =   0
            Width           =   2250
         End
         Begin VB.ComboBox CBOInternalFlag 
            Height          =   315
            ItemData        =   "FrmPO6.frx":03BA
            Left            =   2880
            List            =   "FrmPO6.frx":03BC
            Style           =   2  'Dropdown List
            TabIndex        =   139
            Top             =   315
            Width           =   2730
         End
         Begin VB.TextBox TxtPONo 
            Alignment       =   1  'Right Justify
            Height          =   240
            Left            =   7305
            TabIndex        =   132
            Top             =   -2325
            Visible         =   0   'False
            Width           =   2310
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   4800
            TabIndex        =   121
            Top             =   -300
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.ComboBox CboType 
            Height          =   315
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   2865
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   0
            TabIndex        =   84
            Top             =   210
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   10545
            Locked          =   -1  'True
            TabIndex        =   83
            Top             =   0
            Width           =   1695
         End
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   79
            Top             =   1905
            Width           =   5505
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11385
            TabIndex        =   77
            Top             =   660
            Width           =   825
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11385
            TabIndex        =   76
            Top             =   960
            Width           =   825
         End
         Begin VB.Frame Frame3 
            Caption         =   "»Ì«‰«  «·«⁄ „«œ"
            Height          =   540
            Left            =   -3120
            TabIndex        =   52
            Top             =   2760
            Visible         =   0   'False
            Width           =   3855
            Begin VB.TextBox TxtLcNo 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   600
               TabIndex        =   53
               Top             =   240
               Width           =   2175
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   315
               Left            =   4080
               TabIndex        =   54
               Top             =   600
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   129630209
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   315
               Left            =   4560
               TabIndex        =   55
               Top             =   960
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   129630209
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   315
               Left            =   120
               TabIndex        =   56
               Top             =   960
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   129630209
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker5 
               Height          =   315
               Left            =   4560
               TabIndex        =   57
               Top             =   1320
               Width           =   1620
               _ExtentX        =   2858
               _ExtentY        =   556
               _Version        =   393216
               Format          =   129630209
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker DTPicker6 
               Height          =   315
               Left            =   120
               TabIndex        =   58
               Top             =   1320
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Format          =   129630209
               CurrentDate     =   38784
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷"
               BackColor       =   12632256
               ForeColor       =   16711680
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12632256
               ColorHighlight  =   16777215
               ColorHoverText  =   255
               ColorShadow     =   -2147483637
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledText=   16711680
               ColorToggledHoverText=   255
               ColorTextShadow =   -2147483637
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "„·«ÕŸ« "
               Height          =   375
               Left            =   2400
               TabIndex        =   65
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «·Ê’Ê· «·„ Êﬁ⁄"
               Height          =   255
               Left            =   2280
               TabIndex        =   64
               Top             =   1440
               Width           =   1575
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «· √ŒÌ—"
               Height          =   255
               Left            =   6480
               TabIndex        =   63
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·›⁄·Ì"
               Height          =   375
               Left            =   2640
               TabIndex        =   62
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ «·„ Êﬁ⁄"
               Height          =   375
               Left            =   6480
               TabIndex        =   61
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«· «—ÌŒ"
               Height          =   255
               Left            =   6360
               TabIndex        =   60
               Top             =   600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·«⁄ „«œ"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   59
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1605
            Left            =   2040
            TabIndex        =   39
            Top             =   2970
            Visible         =   0   'False
            Width           =   5685
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               TabIndex        =   42
               Top             =   600
               Width           =   3855
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2640
               TabIndex        =   41
               Top             =   1320
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   240
               TabIndex        =   40
               Top             =   960
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   240
               TabIndex        =   43
               Top             =   1320
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   556
               _Version        =   393216
               Format          =   149684225
               CurrentDate     =   38784
            End
            Begin MSDataListLib.DataCombo DataCombo9 
               Height          =   315
               Left            =   1920
               TabIndex        =   44
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo11 
               Height          =   315
               Left            =   2640
               TabIndex        =   45
               Top             =   960
               Width           =   1425
               _ExtentX        =   2514
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   " «·«‰ Â«¡"
               Height          =   285
               Index           =   24
               Left            =   1680
               TabIndex        =   51
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   23
               Left            =   1560
               TabIndex        =   50
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·Õ”«»"
               Height          =   285
               Index           =   22
               Left            =   4320
               TabIndex        =   49
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„·…"
               Height          =   285
               Index           =   21
               Left            =   4320
               TabIndex        =   48
               Top             =   960
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·»‰ﬂ"
               Height          =   285
               Index           =   20
               Left            =   4320
               TabIndex        =   47
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "‰Ê⁄ «·«„—"
               Height          =   285
               Index           =   19
               Left            =   4320
               TabIndex        =   46
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            Height          =   1500
            Left            =   4560
            TabIndex        =   30
            Top             =   -1485
            Visible         =   0   'False
            Width           =   6600
            Begin VB.CheckBox chkshipped 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «·‘Õ‰"
               Height          =   195
               Left            =   120
               TabIndex        =   80
               Top             =   1320
               Width           =   1815
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   120
               TabIndex        =   31
               Top             =   600
               Width           =   1935
            End
            Begin MSDataListLib.DataCombo DataCombo5 
               Height          =   315
               Left            =   3120
               TabIndex        =   32
               Top             =   1320
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo6 
               Height          =   315
               Left            =   120
               TabIndex        =   33
               Top             =   960
               Width           =   1905
               _ExtentX        =   3360
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo7 
               Height          =   315
               Left            =   3120
               TabIndex        =   34
               Top             =   240
               Width           =   2145
               _ExtentX        =   3784
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DataCombo2 
               Height          =   315
               Left            =   120
               TabIndex        =   74
               Top             =   240
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„‘—Ê⁄"
               Height          =   270
               Index           =   11
               Left            =   2130
               TabIndex        =   75
               Top             =   240
               Width           =   855
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁÌ„…"
               Height          =   285
               Index           =   17
               Left            =   2040
               TabIndex        =   38
               Top             =   600
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "«· ’‰Ì›"
               Height          =   285
               Index           =   16
               Left            =   5400
               TabIndex        =   37
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   285
               Index           =   15
               Left            =   2040
               TabIndex        =   36
               Top             =   960
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ—Ìﬁ… «·‘Õ‰"
               Height          =   285
               Index           =   14
               Left            =   5280
               TabIndex        =   35
               Top             =   1320
               Width           =   1215
            End
         End
         Begin VB.ComboBox CboPriceType 
            Enabled         =   0   'False
            Height          =   315
            Left            =   3570
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   -315
            Visible         =   0   'False
            Width           =   2250
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   10545
            TabIndex        =   0
            Top             =   -210
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2880
            TabIndex        =   21
            Top             =   -390
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   240
            Left            =   1965
            TabIndex        =   20
            Top             =   -345
            Visible         =   0   'False
            Width           =   870
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   270
            Left            =   30
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   -345
            Visible         =   0   'False
            Width           =   1920
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   7305
            TabIndex        =   2
            Top             =   960
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   7305
            TabIndex        =   3
            Top             =   660
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   285
            Left            =   10545
            TabIndex        =   1
            Top             =   315
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   149749761
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   390
            Left            =   7080
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   1305
            Visible         =   0   'False
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmPO6.frx":03BE
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdTemplate 
            Height          =   435
            Left            =   1545
            TabIndex        =   23
            Top             =   1755
            Visible         =   0   'False
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈œ—«Ã ⁄—÷ Ã«Â“"
            BackColor       =   12632256
            ForeColor       =   16711680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
            ColorHighlight  =   16777215
            ColorHoverText  =   255
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   255
            ColorTextShadow =   -2147483637
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   555
            Index           =   4
            Left            =   14745
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1695
            Width           =   3795
            _cx             =   6694
            _cy             =   979
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
            Style           =   1
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
            Begin VB.CheckBox XPChkTAX 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "÷—»Ì»… «·„»Ì⁄« "
               Height          =   330
               Left            =   1860
               TabIndex        =   6
               Top             =   210
               Width           =   1815
            End
            Begin VB.TextBox XPTxtTaxValue 
               Alignment       =   1  'Right Justify
               Height          =   390
               Left            =   30
               TabIndex        =   7
               Top             =   150
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁÌ„…"
               Enabled         =   0   'False
               Height          =   240
               Index           =   4
               Left            =   990
               TabIndex        =   28
               Top             =   285
               Width           =   720
            End
         End
         Begin ImpulseButton.ISButton CmdConvert 
            Height          =   465
            Left            =   1440
            TabIndex        =   66
            Top             =   2970
            Visible         =   0   'False
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕÊÌ· ≈·Ì ›« Ê—…"
            BackColor       =   12632256
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
            ColorHighlight  =   16777215
            ColorHoverText  =   255
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   255
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo Dccurrency 
            Height          =   315
            Left            =   120
            TabIndex        =   72
            Top             =   315
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   7305
            TabIndex        =   81
            Top             =   315
            Width           =   2280
            _ExtentX        =   4022
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   720
            TabIndex        =   122
            Top             =   -315
            Visible         =   0   'False
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo4 
            Height          =   315
            Left            =   7305
            TabIndex        =   134
            Top             =   1275
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCostCenter 
            Height          =   315
            Left            =   120
            TabIndex        =   136
            Top             =   1590
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName1x 
            Height          =   315
            Left            =   13905
            TabIndex        =   147
            Top             =   1590
            Visible         =   0   'False
            Width           =   4065
            _ExtentX        =   7170
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCEquipments 
            Height          =   315
            Left            =   120
            TabIndex        =   152
            Top             =   1275
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpDepartments 
            Height          =   315
            Left            =   7305
            TabIndex        =   156
            Top             =   1590
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   315
            Left            =   7305
            TabIndex        =   159
            Top             =   1905
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName1 
            Height          =   315
            Left            =   120
            TabIndex        =   188
            Top             =   660
            Width           =   4470
            _ExtentX        =   7885
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ› «·ÿ·»"
            Height          =   225
            Index           =   47
            Left            =   12180
            TabIndex        =   191
            Top             =   2220
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ „Œ“‰"
            Height          =   240
            Index           =   46
            Left            =   5850
            TabIndex        =   189
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„œ… «· Ê—Ìœ"
            Height          =   240
            Index           =   11
            Left            =   1800
            TabIndex        =   172
            Top             =   960
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»Ì⁄Â «·‘—«¡"
            Height          =   210
            Index           =   42
            Left            =   5745
            TabIndex        =   163
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·Â «·ÿ·»"
            Height          =   210
            Index           =   41
            Left            =   0
            TabIndex        =   162
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·Â «·ÿ·»"
            Height          =   210
            Index           =   40
            Left            =   3120
            TabIndex        =   161
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ÊŸ›"
            Height          =   255
            Index           =   39
            Left            =   12180
            TabIndex        =   157
            Top             =   1905
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄œÂ/«·”Ì«—…"
            Height          =   255
            Index           =   38
            Left            =   5745
            TabIndex        =   153
            Top             =   1275
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«œ«—… «·ÿ«·»…"
            Height          =   255
            Index           =   37
            Left            =   12180
            TabIndex        =   151
            Top             =   1590
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Œ“‰"
            Height          =   255
            Index           =   36
            Left            =   13425
            TabIndex        =   148
            Top             =   1590
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   210
            Index           =   35
            Left            =   5985
            TabIndex        =   145
            Top             =   0
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·⁄„·Ì…"
            Height          =   165
            Index           =   56
            Left            =   1980
            TabIndex        =   143
            Top             =   30
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·ÿ·»"
            Height          =   255
            Index           =   34
            Left            =   9345
            TabIndex        =   141
            Top             =   30
            Width           =   1095
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   9480
            TabIndex        =   138
            Top             =   315
            Width           =   720
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„—ﬂ“ «· ﬂ·›…"
            Height          =   255
            Index           =   10
            Left            =   5850
            TabIndex        =   137
            Top             =   1590
            Width           =   1095
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»·œ «·„‰‘√"
            Height          =   255
            Index           =   13
            Left            =   12180
            TabIndex        =   135
            Top             =   1275
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   210
            Index           =   33
            Left            =   9525
            TabIndex        =   133
            Top             =   -2325
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‰œÊ»"
            Height          =   240
            Index           =   32
            Left            =   5790
            TabIndex        =   123
            Top             =   -180
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”Ì«”… «·ÿ·»Ì…"
            Height          =   210
            Index           =   18
            Left            =   1800
            TabIndex        =   85
            Top             =   2760
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   11370
            TabIndex        =   82
            Top             =   420
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·“Ê„"
            Height          =   240
            Index           =   28
            Left            =   5880
            TabIndex        =   78
            Top             =   2010
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·Â"
            Height          =   150
            Index           =   12
            Left            =   1875
            TabIndex        =   73
            Top             =   420
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·Â «·ÿ·»"
            Height          =   210
            Index           =   9
            Left            =   5985
            TabIndex        =   29
            Top             =   315
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   240
            Index           =   5
            Left            =   12180
            TabIndex        =   27
            Top             =   30
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ÿ·»"
            Height          =   180
            Index           =   6
            Left            =   12180
            TabIndex        =   26
            Top             =   315
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ «·„Ê’Ï »Â"
            Height          =   210
            Index           =   7
            Left            =   12180
            TabIndex        =   25
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰ «·ÿ«·»"
            Height          =   240
            Index           =   8
            Left            =   12180
            TabIndex        =   24
            Top             =   630
            Width           =   1305
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5835
         Left            =   0
         TabIndex        =   87
         Top             =   2640
         Width           =   14025
         _cx             =   24739
         _cy             =   10292
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
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|Õ«·Â «·«⁄ „«œ"
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
         DogEars         =   0   'False
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   1
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "FrmPO6.frx":0758
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   5370
            Left            =   14670
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   45
            Width           =   13935
            _cx             =   24580
            _cy             =   9472
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   3630
               Left            =   120
               TabIndex        =   120
               Tag             =   "1"
               Top             =   840
               Width           =   13230
               _cx             =   23336
               _cy             =   6403
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPO6.frx":0AF2
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
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
               Height          =   255
               Left            =   9960
               TabIndex        =   130
               Top             =   4560
               Width           =   3375
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5370
            Index           =   15
            Left            =   45
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   45
            Width           =   13935
            _cx             =   24580
            _cy             =   9472
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   12
               Charset         =   178
               Weight          =   700
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
            BorderWidth     =   1
            ChildSpacing    =   1
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5340
               Index           =   16
               Left            =   15
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   15
               Width           =   13905
               _cx             =   24527
               _cy             =   9419
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
               Appearance      =   5
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   8295
                  Index           =   5
                  Left            =   0
                  TabIndex        =   98
                  TabStop         =   0   'False
                  Top             =   -765
                  Width           =   14055
                  _cx             =   24791
                  _cy             =   14631
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
                  Begin VB.TextBox TXTTransactionID4 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   0
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   222
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   1260
                  End
                  Begin VB.TextBox TxtNoteSerial14 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   221
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   1830
                  End
                  Begin VB.CommandButton cmdCreateProduction 
                     Caption         =   "«‰‘«¡ «„— «‰ «Ã"
                     Enabled         =   0   'False
                     Height          =   390
                     Left            =   5700
                     RightToLeft     =   -1  'True
                     TabIndex        =   220
                     Top             =   5400
                     Visible         =   0   'False
                     Width           =   1890
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                     Height          =   1005
                     Left            =   120
                     TabIndex        =   173
                     TabStop         =   0   'False
                     Top             =   1365
                     Width           =   13680
                     _cx             =   24130
                     _cy             =   1773
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
                     Begin VB.TextBox TxtContactPhone 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Left            =   2565
                        TabIndex        =   183
                        Top             =   120
                        Width           =   2235
                     End
                     Begin VB.TextBox Text9 
                        Alignment       =   1  'Right Justify
                        Height          =   330
                        Left            =   11370
                        TabIndex        =   177
                        Top             =   120
                        Width           =   900
                     End
                     Begin VB.TextBox TxtAddress 
                        Alignment       =   1  'Right Justify
                        Height          =   480
                        Left            =   0
                        MultiLine       =   -1  'True
                        ScrollBars      =   2  'Vertical
                        TabIndex        =   176
                        Top             =   480
                        Width           =   4830
                     End
                     Begin VB.TextBox TxtPhone 
                        Alignment       =   1  'Right Justify
                        Height          =   315
                        Left            =   6285
                        TabIndex        =   175
                        Top             =   480
                        Width           =   1035
                     End
                     Begin VB.TextBox TxtCashCustomerName 
                        Alignment       =   1  'Right Justify
                        Height          =   315
                        Left            =   8130
                        TabIndex        =   174
                        Top             =   480
                        Width           =   4125
                     End
                     Begin MSDataListLib.DataCombo DBCboClientName1 
                        Height          =   315
                        Left            =   6270
                        TabIndex        =   178
                        Top             =   120
                        Width           =   4995
                        _ExtentX        =   8811
                        _ExtentY        =   556
                        _Version        =   393216
                        ListField       =   "6"
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSComCtl2.DTPicker DpContactTime 
                        Height          =   270
                        Left            =   0
                        TabIndex        =   184
                        Top             =   150
                        Width           =   1590
                        _ExtentX        =   2805
                        _ExtentY        =   476
                        _Version        =   393216
                        CustomFormat    =   "'Time: 'hh:mm tt"
                        Format          =   151977987
                        UpDown          =   -1  'True
                        CurrentDate     =   40909
                     End
                     Begin VB.Label Label17 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "—ﬁ„ «·« ’«·"
                        ForeColor       =   &H00000000&
                        Height          =   270
                        Left            =   4800
                        TabIndex        =   186
                        Top             =   150
                        Width           =   1320
                     End
                     Begin VB.Label Label16 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Êﬁ  «·« ’«·"
                        ForeColor       =   &H00000000&
                        Height          =   270
                        Left            =   1215
                        TabIndex        =   185
                        Top             =   150
                        Width           =   1305
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·⁄„Ì· "
                        Height          =   240
                        Index           =   45
                        Left            =   12420
                        TabIndex        =   182
                        Top             =   120
                        Width           =   1065
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·⁄‰Ê«‰"
                        Height          =   270
                        Index           =   44
                        Left            =   4830
                        TabIndex        =   181
                        Top             =   630
                        Width           =   960
                     End
                     Begin VB.Label Label15 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " ·Ì›Ê‰"
                        ForeColor       =   &H00000000&
                        Height          =   255
                        Left            =   7440
                        TabIndex        =   180
                        Top             =   555
                        Width           =   570
                     End
                     Begin VB.Label Label14 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
                        ForeColor       =   &H00000000&
                        Height          =   255
                        Left            =   12300
                        TabIndex        =   179
                        Top             =   555
                        Width           =   1320
                     End
                  End
                  Begin VB.Frame Frame4 
                     BorderStyle     =   0  'None
                     Height          =   900
                     Left            =   6525
                     TabIndex        =   99
                     Top             =   5835
                     Visible         =   0   'False
                     Width           =   1710
                     Begin DBPIXLib.DBPix20 DBPix202 
                        Height          =   855
                        Left            =   240
                        TabIndex        =   100
                        Top             =   120
                        Width           =   2415
                        _Version        =   131072
                        _ExtentX        =   4260
                        _ExtentY        =   1508
                        _StockProps     =   1
                        _Image          =   "FrmPO6.frx":0C35
                        ImageResampleWidth=   100
                        ImageResampleHeight=   100
                        ImageResampleMode=   1
                        ImageSaveFormat =   0
                        JPEGQuality     =   75
                        JPEGEncoding    =   0
                        JPEGColorMode   =   0
                        JPEGNoRecompress=   -1  'True
                        JPEGRotateWarning=   0
                        PNGColorDepth   =   0
                        PNGCompression  =   0
                        PNGFilter       =   0
                        PNGInterlace    =   1
                        ImageDitherMethod=   3
                        ImagePaletteMethod=   4
                        ImagePreviewMode=   0   'False
                        ImageKeepMetaData=   -1  'True
                        UseAmbientBackcolor=   -1  'True
                        ViewAsyncDecoding=   -1  'True
                        ViewEnableMouseZoom=   -1  'True
                        ViewInitialZoom =   1
                        ViewHAlign      =   1
                        ViewVAlign      =   1
                        ViewMenuMode    =   0
                     End
                     Begin VB.Label LblPostedPerson 
                        Alignment       =   2  'Center
                        BackStyle       =   0  'Transparent
                        Caption         =   "."
                        Height          =   255
                        Left            =   3600
                        TabIndex        =   103
                        Top             =   240
                        Width           =   1695
                     End
                     Begin VB.Label Label10 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·‰ÊﬁÌ⁄"
                        Height          =   255
                        Left            =   2640
                        TabIndex        =   102
                        Top             =   240
                        Width           =   855
                     End
                     Begin VB.Label Label4 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Ì⁄ „œ"
                        Height          =   255
                        Left            =   5160
                        TabIndex        =   101
                        Top             =   240
                        Width           =   735
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   1200
                     Index           =   2
                     Left            =   0
                     TabIndex        =   104
                     TabStop         =   0   'False
                     Top             =   2490
                     Width           =   13605
                     _cx             =   23998
                     _cy             =   2117
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
                     Begin VB.ComboBox CboItemCase 
                        Height          =   315
                        Left            =   3270
                        Style           =   2  'Dropdown List
                        TabIndex        =   107
                        Top             =   300
                        Width           =   1560
                     End
                     Begin VB.TextBox TxtQuantity 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Left            =   2160
                        MaxLength       =   10
                        TabIndex        =   106
                        Top             =   300
                        Width           =   1110
                     End
                     Begin VB.TextBox TxtPrice 
                        Alignment       =   1  'Right Justify
                        Height          =   300
                        Left            =   780
                        MaxLength       =   10
                        TabIndex        =   105
                        Top             =   300
                        Width           =   1365
                     End
                     Begin MSDataListLib.DataCombo DCboItemsName 
                        Height          =   315
                        Left            =   4815
                        TabIndex        =   108
                        Top             =   300
                        Width           =   2745
                        _ExtentX        =   4842
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin MSDataListLib.DataCombo DCboItemsCode 
                        Height          =   315
                        Left            =   7560
                        TabIndex        =   109
                        Top             =   300
                        Width           =   2025
                        _ExtentX        =   3572
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin ImpulseButton.ISButton CmdAdd 
                        Height          =   375
                        Left            =   75
                        TabIndex        =   110
                        Top             =   270
                        Width           =   630
                        _ExtentX        =   1111
                        _ExtentY        =   661
                        ButtonStyle     =   1
                        ButtonPositionImage=   4
                        Caption         =   ""
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
                        ButtonImage     =   "FrmPO6.frx":0C4D
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
                     Begin MSDataListLib.DataCombo DCPROJECT1 
                        Height          =   315
                        Left            =   10260
                        TabIndex        =   164
                        Top             =   840
                        Width           =   3195
                        _ExtentX        =   5636
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
                     Begin MSDataListLib.DataCombo Dcterm1 
                        Height          =   315
                        Left            =   6945
                        TabIndex        =   165
                        Top             =   840
                        Width           =   3255
                        _ExtentX        =   5741
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
                     Begin MSDataListLib.DataCombo dcopr 
                        Height          =   315
                        Left            =   5040
                        TabIndex        =   166
                        Top             =   840
                        Width           =   1890
                        _ExtentX        =   3334
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
                     Begin MSDataListLib.DataCombo XPCboGroupBuiltin 
                        Height          =   315
                        Left            =   9690
                        TabIndex        =   218
                        Top             =   270
                        Width           =   2535
                        _ExtentX        =   4471
                        _ExtentY        =   556
                        _Version        =   393216
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·„Ã„Ê⁄…"
                        Height          =   360
                        Index           =   77
                        Left            =   12330
                        TabIndex        =   219
                        Top             =   210
                        Width           =   765
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   " «·„‘—Ê⁄"
                        Height          =   315
                        Index           =   48
                        Left            =   11685
                        TabIndex        =   169
                        Top             =   600
                        Width           =   720
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·»‰œ"
                        Height          =   315
                        Index           =   43
                        Left            =   8325
                        TabIndex        =   168
                        Top             =   600
                        Width           =   720
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·⁄„·Ì…"
                        Height          =   315
                        Index           =   51
                        Left            =   5445
                        TabIndex        =   167
                        Top             =   600
                        Width           =   750
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "ﬂÊœ «·’‰›"
                        Height          =   255
                        Index           =   31
                        Left            =   7020
                        TabIndex        =   115
                        Top             =   0
                        Width           =   3015
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈”„ «·’‰›"
                        Height          =   255
                        Index           =   30
                        Left            =   4500
                        TabIndex        =   114
                        Top             =   0
                        Width           =   3000
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "Õ«·… «·’‰›"
                        Height          =   255
                        Index           =   29
                        Left            =   3060
                        TabIndex        =   113
                        Top             =   0
                        Width           =   1680
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·ﬂ„Ì…"
                        Height          =   255
                        Index           =   27
                        Left            =   1740
                        TabIndex        =   112
                        Top             =   0
                        Width           =   1890
                     End
                     Begin VB.Label lbl 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·”⁄—"
                        Height          =   255
                        Index           =   26
                        Left            =   345
                        TabIndex        =   111
                        Top             =   0
                        Width           =   1950
                     End
                  End
                  Begin VSFlex8UCtl.VSFlexGrid FG 
                     Height          =   1665
                     Left            =   150
                     TabIndex        =   116
                     Top             =   3735
                     Width           =   13680
                     _cx             =   24130
                     _cy             =   2937
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
                     Rows            =   2
                     Cols            =   35
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   300
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   -1  'True
                     FormatString    =   $"FrmPO6.frx":0FE7
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
                     WallPaperAlignment=   0
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   24
                  End
                  Begin MSComctlLib.Toolbar TBar 
                     Height          =   630
                     Left            =   510
                     TabIndex        =   117
                     Top             =   5415
                     Width           =   3210
                     _ExtentX        =   5662
                     _ExtentY        =   1111
                     ButtonWidth     =   609
                     ButtonHeight    =   1005
                     Appearance      =   1
                     _Version        =   393216
                  End
                  Begin ImpulseButton.ISButton Accredit 
                     Height          =   540
                     Left            =   3750
                     TabIndex        =   131
                     Top             =   5340
                     Width           =   1860
                     _ExtentX        =   3281
                     _ExtentY        =   953
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "«—”«· ··«⁄ „«œ"
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
                  Begin VB.Label LblItemsCount 
                     Alignment       =   2  'Center
                     BackColor       =   &H00404040&
                     ForeColor       =   &H0000FFFF&
                     Height          =   285
                     Left            =   30
                     TabIndex        =   118
                     Top             =   5415
                     Width           =   450
                  End
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Label12"
                  Height          =   900
                  Left            =   2895
                  TabIndex        =   97
                  Top             =   255
                  Width           =   960
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   3165
                  Index           =   62
                  Left            =   2805
                  TabIndex        =   90
                  Top             =   1410
                  Width           =   510
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   5340
               Index           =   9
               Left            =   15
               TabIndex        =   91
               TabStop         =   0   'False
               Top             =   15
               Width           =   13905
               _cx             =   24527
               _cy             =   9419
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
               Appearance      =   5
               MousePointer    =   0
               Version         =   801
               BackColor       =   14871017
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   0
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
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   2775
                  Left            =   4695
                  TabIndex        =   93
                  Top             =   1410
                  Width           =   945
               End
               Begin VB.TextBox Text8 
                  Alignment       =   1  'Right Justify
                  Height          =   4230
                  Left            =   3585
                  MaxLength       =   4
                  TabIndex        =   92
                  Top             =   915
                  Width           =   630
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3225
                  Index           =   69
                  Left            =   3315
                  TabIndex        =   96
                  Top             =   1410
                  Width           =   270
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   2685
                  Index           =   68
                  Left            =   4215
                  TabIndex        =   95
                  Top             =   1740
                  Width           =   300
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   2775
                  Index           =   67
                  Left            =   2805
                  TabIndex        =   94
                  Top             =   1410
                  Width           =   510
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   615
         Index           =   6
         Left            =   0
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   0
         Width           =   14010
         _cx             =   24712
         _cy             =   1085
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
         Caption         =   "ÿ·»«  œ«Œ·Ì…"
         Align           =   0
         AutoSizeChildren=   7
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
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3150
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   720
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   5970
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   0
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   194
            Top             =   0
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   4785
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   193
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2220
            TabIndex        =   197
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
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
            ButtonImage     =   "FrmPO6.frx":158E
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
            Left            =   1110
            TabIndex        =   198
            Top             =   30
            Width           =   1065
            _ExtentX        =   1879
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
            ButtonImage     =   "FrmPO6.frx":1928
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
            Left            =   3150
            TabIndex        =   199
            Top             =   30
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "FrmPO6.frx":1CC2
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
            Left            =   60
            TabIndex        =   200
            Top             =   30
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "FrmPO6.frx":205C
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   345
            Left            =   8895
            TabIndex        =   201
            Top             =   120
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmPO6.frx":23F6
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   345
            Left            =   9945
            TabIndex        =   202
            Top             =   120
            Visible         =   0   'False
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmPO6.frx":2790
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   480
            Left            =   7260
            TabIndex        =   203
            Top             =   0
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   847
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmPO6.frx":2D2A
            ButtonImageHover=   "FrmPO6.frx":3A04
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Image ImgFavorites 
            Height          =   390
            Left            =   9585
            Picture         =   "FrmPO6.frx":46DE
            Stretch         =   -1  'True
            Top             =   0
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
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
            Height          =   435
            Index           =   52
            Left            =   4230
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   120
            Width           =   7890
         End
         Begin VB.Label LblShortcutKeys 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "ÃœÌœ F12 Or Enter ,  ⁄œÌ· F11 , Õ›Ÿ F10 ,  —«Ã⁄ F9 ,Õ–› F8 ,»ÕÀ F7 "
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
            Height          =   195
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   390
            Visible         =   0   'False
            Width           =   8250
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   555
         Index           =   1
         Left            =   0
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   9120
         Width           =   16080
         _cx             =   28363
         _cy             =   979
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
            Height          =   390
            Index           =   0
            Left            =   12090
            TabIndex        =   207
            Top             =   90
            Width           =   930
            _ExtentX        =   1640
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
            Left            =   10650
            TabIndex        =   208
            Top             =   90
            Width           =   945
            _ExtentX        =   1667
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
            Left            =   8775
            TabIndex        =   209
            Top             =   90
            Width           =   1050
            _ExtentX        =   1852
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
            Height          =   390
            Index           =   3
            Left            =   7410
            TabIndex        =   210
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
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
            Left            =   6240
            TabIndex        =   211
            Top             =   90
            Width           =   915
            _ExtentX        =   1614
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
            Height          =   390
            Index           =   5
            Left            =   5085
            TabIndex        =   212
            Top             =   90
            Width           =   975
            _ExtentX        =   1720
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   6
            Left            =   1950
            TabIndex        =   213
            Top             =   90
            Width           =   900
            _ExtentX        =   1588
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
            Index           =   7
            Left            =   3960
            TabIndex        =   214
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄…"
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
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   390
            Left            =   2985
            TabIndex        =   215
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„”«⁄œ…"
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   390
            Left            =   360
            TabIndex        =   216
            Top             =   120
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   688
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
            Index           =   8
            Left            =   -360
            TabIndex        =   217
            Top             =   90
            Visible         =   0   'False
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â ÿ·» ‘—«¡ "
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
      End
   End
End
Attribute VB_Name = "FrmPO6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim NewGrid As New ClsGrid
Dim SaleReport As ClsSaleReport
Dim cSearchDcbo(3)   As clsDCboSearch
 Dim CurrentTransactionType As Integer


Private Sub Dcbranch_Change()
 Dim Dcombos As ClsDataCombos
        Set Dcombos = New ClsDataCombos

  If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
TxtNoteSerial1.Text = ""
     Dcombos.GetStores Me.DCboStoreName, val(dcBranch.BoundText)
     Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True, val(dcBranch.BoundText)
    Dcombos.GetEmployees Me.DcboEmpName, , , val(dcBranch.BoundText)
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName1, True, val(dcBranch.BoundText)
End If
End Sub
Private Sub CBoBasedON_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then

    If TXT_order_no.Text <> "" Then
        TXT_order_no.Text = ""
    End If
    
  End If
End Sub

Private Sub CBoBasedON_Click()
CBoBasedON_Change
End Sub

Private Sub CBOOrderType_Change()

If CBOOrderType.ListIndex = 1 Then
lbl(36).Visible = True
'TxtStoreID1.Visible = True
'DCboStoreName1.Visible = True
Else
lbl(36).Visible = False
'TxtStoreID1.Visible = False
'DCboStoreName1.Visible = False
End If

End Sub

Private Sub CBOOrderType_Click()
CBOOrderType_Change
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim intDef As Integer
  On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = val(GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2))
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
            DCboStoreName.BoundText = intDef
            Dccurrency.BoundText = 1
'            Fg.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            Me.CboPriceType.ListIndex = 0
            CBOInternalFlag.ListIndex = 0
                     GRID2.Rows = 1
            
Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = True
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
                    
                    
        

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.dcBranch.Enabled = True
                                   Else
                                    Me.dcBranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = True
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If

            
            
            
            Me.dcBranch.BoundText = Current_branch
            DBPix202.ImageClear
Accredit.Enabled = True
  DcCostCenter.Text = ""
                If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               
         CBOOrderType.ListIndex = 0
          DCOPrType.ListIndex = 0
           CBoBasedON.ListIndex = 0
           
        Case 1


            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If


              If ChekClodePeriod(Me.XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "› —Â „€·ﬁ… "
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If


            If ScreenAproved(val(Me.XPTxtBillID.Text), Me.Name) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«·Õ—ﬂÂ „— »ÿÂ »«·«⁄ „«œ« "
                Else
                    MsgBox "Can not edit.This process associated with approvals"
                End If
                Exit Sub
            End If
            
            
            TxtModFlg.Text = "E"
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id
        cmdCreateProduction.Enabled = False
        Case 2
            Dim Msg  As String

            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If
              
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBuySearch.DealingForm = GridTransType.internalorder
'If SystemOptions.UserInterface = ArabicInterface Then
'FrmBuySearch.XPLbl(0).Caption = "«·„Ê—œ «·„Ê’Ì"
'  With Me.Fg
'
'      .TextMatrix(0, .ColIndex("StorName")) = "Store Name"
'    End With

            
            
           FrmBuySearch.Index = 0
              FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ÿ·»«   œ«Œ·Ì…"
             FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
If CBOOrderType.ListIndex <> 3 Then
            PrintReport
Else
print_report
End If
'print_report

        Case 8
            On Error GoTo ErrTrap

            If XPTxtBillID.Text <> "" Then
                Set SaleReport = New ClsSaleReport
                SaleReport.ShowPrice XPTxtBillID.Text, 6, DcboEmpName.Text, val(DBCboClientName.BoundText)
            End If

            '        PrintReport1 (Txt_order_no.text)
        Case 6
            Unload Me
    End Select

    Exit Sub
ErrTrap:
End Sub

Function PrintReport1(order_no As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

    MySQL = "Select * From QRY_items_orders_data where order_no='" & order_no & "'"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Order_status.rpt"
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
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = "Order status"
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, ""

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function

Private Sub cmdAdd_Click()
AddNewFgAttachRow
End Sub
Private Sub AddNewFgAttachRow()
    Dim Msg As String
    Dim LngFindRow As Long
    Dim LngNewRow As Long
    LngNewRow = ModFgLib.SetFgForNewRow(FG, FG.ColIndex("Name"))
    With Me.FG
        .TextMatrix(LngNewRow, .ColIndex("projectid")) = Me.DCPROJECT1.BoundText
        .TextMatrix(LngNewRow, .ColIndex("project")) = DCPROJECT1.Text
        .TextMatrix(LngNewRow, .ColIndex("pandid")) = Dcterm1.BoundText
        .TextMatrix(LngNewRow, .ColIndex("pand")) = Dcterm1.Text
        .TextMatrix(LngNewRow, .ColIndex("proid")) = dcopr.BoundText
        .TextMatrix(LngNewRow, .ColIndex("pro")) = dcopr.Text
        
        .TextMatrix(LngNewRow, .ColIndex("GroupID")) = XPCboGroupBuiltin.BoundText
        .TextMatrix(LngNewRow, .ColIndex("GroupName")) = XPCboGroupBuiltin.Text
        
    
       ' .AutoSize 0, .Cols - 1, False
    End With


End Sub
Private Sub CmdConvert_Click()
    Dim RowNum As Integer
    Dim Frm As Form
    On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.CboPriceType.ListIndex = 0 Then
        Set Frm = New frmsalebill
    ElseIf Me.CboPriceType.ListIndex = 1 Then
        Set Frm = New FrmBillBuy
    End If

    With Frm
        .Convert
        '    .XPTxtBillID.Text = XPTxtBillID.Text
        .XPDtbBill.value = XPDtbBill.value
        .DBCboClientName.BoundText = DBCboClientName.BoundText
        .DCboStoreName.BoundText = DCboStoreName.BoundText
        .Dccurrency.BoundText = Me.Dccurrency.BoundText

        For RowNum = 1 To FG.Rows - 1

            If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
                .FG.Rows = .FG.Rows + 1
            End If

            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
            .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
            Dim StrSQL As String
            Dim RsUnit As New ADODB.Recordset
        
            StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 6) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
            .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(ISNULL(RsUnit("UnitID")), "", (RsUnit("UnitID").value))
            .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = IIf(ISNULL(RsUnit("UnitName")), "", (RsUnit("UnitName").value))
        Next RowNum

        .Cala
    End With

    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdTemplate_Click()
    Dim Frm  As FrmBuySearch
    On Error GoTo ErrTrap
    Set Frm = New FrmBuySearch

    With Frm
        .DealingForm = InsertTemplate
        .Caption = "«·⁄—Ê÷ «·Ã«Â“…"
        '    .MDIChild = True
        .BorderStyle = 0
        '  .MinButton = True
        .show vbModeless, mdifrmmain
        .Visible = True
    End With

    Exit Sub
ErrTrap:
End Sub

 

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
      FrmCompanySearch.lblSearchtype.Caption = 5
        FrmCompanySearch.show vbModal
        
        
    End If
          
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos

       
            Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
      
    End If

End Sub
 
Private Sub DcboEmp_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

 

End Sub

Private Sub DcboEmpName_Change()
 'If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
         If val(Me.DcboEmpName.BoundText) = 0 Then Exit Sub
           Me.txtempcode.Text = get_EMPLOYEE_Data(val(Me.DcboEmpName.BoundText), "Fullcode")
        'DCEmP.text = DCEmP.text
'End If

End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.Text = DCboItemsCode.Text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 200
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboItemsName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF9 Then
                    
        FrmSearchSerial.XPTxtCode.Text = DCboItemsCode.Text
        FrmSearchSerial.show
        FrmSearchSerial.Cmd_Click (0)
                    
    End If

    If KeyCode = vbKeyF3 Then
        
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 200
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 38) = True Then
    ' TxtNoteSerial.text = ""
    TxtNoteSerial1.Text = ""

     End If
     
    End If
End Sub

Private Sub DCboStoreName_Click(Area As Integer)
'DCboStoreName_Change
End Sub

Private Sub DCboStoreName_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetStores Me.DCboStoreName

    End If
        
End Sub

Private Sub DCboStoreName1_Change()
 TxtStoreID1.Text = getStoreCoding(val(DCboStoreName1.BoundText))
  NewGrid.mStorePurName = DCboStoreName1.Text
   NewGrid.DefStorePurchase = val(DCboStoreName1.BoundText)
End Sub

Private Sub Dcbranch_Click(Area As Integer)
'dcBranch_Change

End Sub

 

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
    End If

End Sub

Private Sub DcCostCenter_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        CostCenterSearch.show
        CostCenterSearch.RetrunType = 10
    End If

    If KeyCode = vbKeyF5 Then
        Dim StrSQL As String
        StrSQL = "  SELECT code ,account_name FROM markaas_taklefa  WHERE level=3 and NOT(account_no IS NULL)  "
        fill_combo Me.DcCostCenter, StrSQL
    End If
    
End Sub

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim My_SQL As String
        My_SQL = " select id,code from currency"
 
        fill_combo Me.Dccurrency, My_SQL

    End If

End Sub
Private Sub dcproject1_Click(Area As Integer)
If val(DCPROJECT1.BoundText) <> 0 Then
fillterms val(DCPROJECT1.BoundText)
End If
End Sub
Function fillterms(project_id As Integer)
    Dim My_SQL As String
 
    My_SQL = " select oprid,des from dbo.projects_des where project_id=" & project_id

    fill_combo Me.Dcterm1, My_SQL
     
    Dcterm1.ReFill
End Function
Private Sub Dcterm1_Click(Area As Integer)
 Dim Dcombos As ClsDataCombos

       Set Dcombos = New ClsDataCombos
  If DCPROJECT1.BoundText <> "" Then
        
         If Me.Dcterm1.BoundText <> "" Then
         Dcombos.GetProcessOfProjedt dcopr, val(DCPROJECT1.BoundText), , val(Dcterm1.BoundText), 2
         End If
       
    End If
End Sub
Private Sub Ele_Click(Index As Integer)

    Select Case Index

        Case 6
            On Error GoTo ErrTrap
            '        If Me.WindowState = vbNormal Then
            '            Me.WindowState = vbMaximized
            '        Else
            '            Me.WindowState = vbNormal
            '        End If
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

MySQL = "SELECT  Transaction_Details.RecivePeriod ,  dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID, dbo.Transaction_Details.ItemDiscountType, "
MySQL = MySQL & "                      dbo.Transaction_Details.ItemDiscount, dbo.Transactions.order_no, dbo.Transactions.Currency_id, dbo.Transaction_Details.Item_ID,"
MySQL = MySQL & "                      dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ColorID,"
MySQL = MySQL & "                      dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ClassId, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
MySQL = MySQL & "                      dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblItemsclasses.SizeName AS ClassName,"
MySQL = MySQL & "                      dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Transaction_Type, dbo.TblCustemers.Fullcode,"
MySQL = MySQL & "                     dbo.TblCustemers.E_mail, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.FaxNumber, dbo.Transaction_Details.ParrtNoCode,"
MySQL = MySQL & "                      dbo.TblUnites.UnitNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.DepartementID,"
MySQL = MySQL & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price,"
MySQL = MySQL & "                      dbo.Transaction_Details.RequestLimit, dbo.Transaction_Details.LastPurchasePrice, dbo.Transaction_Details.LastPurchaseqty,"
MySQL = MySQL & "                      dbo.Transactions.purchaseType, dbo.Transaction_Details.AverageIssue, dbo.Transaction_Details.AverageIssueyraly, dbo.Transaction_Details.LastPurchaseDate, dbo.Transactions.InternalFlag, dbo.Transactions.FixesAssetsID,"
MySQL = MySQL & "                      dbo.FixedAssets.code AS Fixedcode, dbo.FixedAssets.Name, dbo.FixedAssets.namee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
MySQL = MySQL & "                      dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblItems.Fullcode AS ItemFullcode, dbo.Transaction_Details.ItemSerial,"
MySQL = MySQL & "                      dbo.Transaction_Details.ItemCase, dbo.markaas_taklefa.account_name, dbo.markaas_taklefa.Code AS TaklfaCode, dbo.Transaction_Details.ItemBalance,"
MySQL = MySQL & "                      dbo.Transactions.Priod , dbo.Transactions.PriodDMY,Shipping_Pos"
MySQL = MySQL & " FROM         dbo.Transactions INNER JOIN"
MySQL = MySQL & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
MySQL = MySQL & "                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
MySQL = MySQL & "                      dbo.TblItemsclasses ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.markaas_taklefa ON dbo.Transactions.general_cost_center = dbo.markaas_taklefa.Code LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.Transactions.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmpDepartments ON dbo.Transactions.DepartementID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItemsSizes ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
MySQL = MySQL & "                     dbo.TblItemsColors ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"

MySQL = MySQL & "  Where (dbo.Transactions.Transaction_ID =" & val(XPTxtBillID.Text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
   '      StrFileName = App.path & "\Reports\REPORTS NEW\PerformaInvoices777Sh.rpt"
         StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PerformaInvoices777Sh.rpt"
         
         
     Else
     '  StrFileName = App.path & "\Reports\REPORTS NEW\PerformaInvoices777Sh.rpt"
     StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\PerformaInvoices777Sh.rpt"
     
       
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
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
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function

Private Sub Fg_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , , Me.TXT_order_no
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), , Me.TXT_order_no

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub FG_CellButtonClick(ByVal Row As Long, _
                               ByVal Col As Long)

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    FrmAddNewItem.Tag = "xx"
'        FrmAddNewItem.DealingForm = ShowPrice
'        FrmAddNewItem.show vbModal

    End If
Select Case FG.ColKey(Col)
Case "NoteSerial14"
            TXTTransactionID4 = val(FG.TextMatrix(Row, FG.ColIndex("TransactionID4")))
            FrmProductionOrder.show
         FrmProductionOrder.XPBtnMove_Click (2)
        FrmProductionOrder.Retrive val(TXTTransactionID4.Text)


    
    Case "ShowAttatch2"
        If Trim(FG.TextMatrix(Row, FG.ColIndex("ShowAttatch"))) = "" Then
            FG.TextMatrix(Row, FG.ColIndex("ShowAttatch")) = Trim(FG.TextMatrix(Row, FG.ColIndex("Code"))) & Row & XPTxtBillID & user_id
        End If
        Dim mItemFullCode As String
            mItemFullCode = Trim(FG.TextMatrix(Row, FG.ColIndex("ShowAttatch")))
            
            On Error Resume Next
            ShowAttachments mItemFullCode, "0741201407"
    End Select


End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If TxtModFlg = "R" Then
    Select Case FG.ColKey(Col)
    Case "ShowAttatch2"
        FG.EditMaxLength = 10
    Case Else
        Cancel = True
    End Select
End If
End Sub

Private Sub Form_Activate()
    'XPTxtBillID.SetFocus
End Sub

 

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLcNo.Text)
    'Frame3.Visible = True
End Sub

Private Sub ISButton2_Click()
    On Error Resume Next
ShowAttachments TxtNoteSerial1, "310319"
 
End Sub

Private Sub Label10_Click()
    Frame3.Visible = False
End Sub
 
Private Sub Accredit_Click()
    Dim sql As String
    Dim BeginTrans As Boolean
    'sql = "update  Transactions  set Posted=" & user_id & "  where Transaction_ID=" & Val(XPTxtBillID.text)
    'Cn.Execute sql

    Cn.BeginTrans
    BeginTrans = True

    If ISNULL(rs("Posted")) Then
        rs("Posted") = user_id
        rs("PostedDate") = Time
    Else
        rs("Posted") = Null
       rs("PostedDate") = Time
    End If
   
    rs.update
 If SystemOptions.UserInterface = ArabicInterface Then
    Accredit.Caption = " „ «·«—”«· ··«⁄ „«œ"
Else
Accredit.Caption = "Sent To approval "
End If

    Cn.CommitTrans
    BeginTrans = False
FillApprovedTable
    Retrive (val(XPTxtBillID.Text))

End Sub
Function FillApprovedTable()
 Dim RSApproval  As New ADODB.Recordset
   Set RSApproval = New ADODB.Recordset
   Dim currentdate As Date
   RSApproval.Open "[ApprovalData]", Cn, adOpenStatic, adLockOptimistic, adCmdTable


 Dim sql As String
  Dim Rs1 As New ADODB.Recordset
 Dim i As Integer
    sql = "SELECT     TOP 100 PERCENT dbo.TblApprovalDef.ScreenName, dbo.TblApprovalDefDetails.PlainMessageID AS levelo, dbo.TbllevelWorker.EmpID, "
  sql = sql & " dbo.TblApprovalDefDetails.id AS levelorder, dbo.TbllevelWorker.id AS currorder"
  sql = sql & " FROM         dbo.TblApprovalDef INNER JOIN"
  sql = sql & " dbo.TblApprovalDefDetails ON dbo.TblApprovalDef.id = dbo.TblApprovalDefDetails.lMessageDefID INNER JOIN"
  sql = sql & "  dbo.TbllevelWorker ON dbo.TblApprovalDefDetails.PlainMessageID = dbo.TbllevelWorker.LevelID"
sql = sql & " WHERE     (dbo.TblApprovalDef.ScreenName = N'" & Me.Name & "')"
sql = sql & " and  ( dbo.TblApprovalDef.BranchId =0 or     dbo.TblApprovalDef.BranchId =" & val(Me.dcBranch.BoundText) & ")"
sql = sql & " ORDER BY dbo.TblApprovalDefDetails.id, dbo.TbllevelWorker.id  "

    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
Dim UserID As Integer
Dim EmpID As Integer
    If Rs1.RecordCount > 0 Then
            currentdate = Now
            
                                    GetApprovalDepartement val(DcboEmpDepartments.BoundText), UserID, EmpID
            
            If UserID <> 0 Then
           '***************************************
                                 RSApproval.AddNew
                        RSApproval("ScreenName").value = Me.Name
                        RSApproval("levelo").value = 1
                       RSApproval("EmpID").value = UserID
                        RSApproval("levelorder").value = 1
                         RSApproval("currorder").value = 1
                          RSApproval("Transaction_ID").value = val(XPTxtBillID.Text)
                          RSApproval("NoteSerial").value = TxtNoteSerial1.Text
                        RSApproval("Transaction_Date").value = Date
                        
                          RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
                       RSApproval("SendTime").value = currentdate
        
                 
                                RSApproval("Currcursor").value = 1
                                 RSApproval("FromUser").value = user_name
                     
                        
                        RSApproval.update
              End If
              
            For i = 1 To Rs1.RecordCount
              RSApproval.AddNew
                RSApproval("ScreenName").value = Me.Name
                RSApproval("levelo").value = IIf(ISNULL(Rs1("levelo").value), Null, Rs1("levelo").value)
               RSApproval("EmpID").value = IIf(ISNULL(Rs1("EmpID").value), Null, Rs1("EmpID").value)
                RSApproval("levelorder").value = IIf(ISNULL(Rs1("levelorder").value), Null, Rs1("levelorder").value)
                 RSApproval("currorder").value = IIf(ISNULL(Rs1("currorder").value), Null, Rs1("currorder").value)
                  RSApproval("Transaction_ID").value = val(XPTxtBillID.Text)
                  RSApproval("NoteSerial").value = TxtNoteSerial1.Text
                RSApproval("Transaction_Date").value = Date
                
                  RSApproval("ExpectedtimeTime").value = DateAdd("N", GetTimeforTransaction(Me.Name), currentdate)
               RSApproval("SendTime").value = currentdate

                 If i = 1 And UserID = 0 Then
                        RSApproval("Currcursor").value = 1
                         RSApproval("FromUser").value = user_name
                End If
                
                RSApproval.update
                Rs1.MoveNext
            Next i

    End If
    
    

End Function
Public Sub RetriveOrder(Optional order_no As String = "", _
                        Optional Transaction_Type As Integer = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

   If Transaction_Type <> 20 Then
        StrSQL = "Select * from transactions where  Transaction_Type=" & Transaction_Type & " and Order_no='" & order_no & "'"
 Else
         StrSQL = "Select * from transactions where  Transaction_Type=" & Transaction_Type & " and NoteSerial1='" & order_no & "'"

 End If
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Sub
    Else
        DBCboClientName.BoundText = IIf(ISNULL(rs("CusID").value), "", rs("CusID").value)
        Me.Dccurrency.BoundText = IIf(ISNULL(rs("Currency_id").value), "", rs("Currency_id").value)
        Me.DCboStoreName.BoundText = IIf(ISNULL(rs("storeid").value), "", rs("storeid").value)
        Me.dcBranch.BoundText = IIf(ISNULL(rs("Branchid").value), "", rs("Branchid").value)

        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass

    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(ISNULL(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(ISNULL(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(ISNULL(RsDetails("showqty")), "", (RsDetails("showqty").value))

            'FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").value))
            If Transaction_Type = 0 Or Transaction_Type = 20 Then
                FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(ISNULL(RsDetails("ShowPrice")), 0, (RsDetails("ShowPrice").value)) ' GET_COST_PRICE_FOR_PRODUCT_ITEM(Val(FG.TextMatrix(Num, FG.ColIndex("Code"))))
            End If
      
            '  FG.TextMatrix(Num, FG.ColIndex("Expenses")) = IIf(IsNull(RsDetails("Lineexpenses")), "", (RsDetails("Lineexpenses").value))
         
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(ISNULL(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(ISNULL(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(ISNULL(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(ISNULL(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(ISNULL(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(ISNULL(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemType")) = IIf(ISNULL(RsDetails("ItemType")), 0, (RsDetails("ItemType").value))
         
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(ISNULL(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(ISNULL(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    TxtFillData.Text = "F"
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub Txt_order_no_Change()
 
    Dim Transaction_Type As Integer
    If CBoBasedON.ListIndex = 1 Then
        Transaction_Type = 6
   ElseIf CBoBasedON.ListIndex = 2 Then
 
        
    ElseIf CBoBasedON.ListIndex = 3 Then
    
   
       ElseIf CBoBasedON.ListIndex = 5 Then

        Transaction_Type = 20
 
    Else
     
         Exit Sub
    End If

   ' Transaction_ID = get_transactionData("order_no", Txt_order_no.text, "Transaction_ID", Transaction_Type)
'


    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TXT_order_no, Transaction_Type
    End If
 



End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, Shift As Integer)


Dim transactiontype As Integer
Dim transactionName As String

    If KeyCode = vbKeyF3 Then
        
       If CBoBasedON.ListIndex = 1 Then
        transactiontype = 6
                      If SystemOptions.UserInterface = ArabicInterface Then
                          transactionName = "»ÕÀ ⁄‰ «Ê«„— «·»Ì⁄"
                        Else
                        transactionName = "Search  Sales Order"
                        End If
      
        ElseIf CBoBasedON.ListIndex = 2 Then

                        
        Else
    '    transactiontype = 0
        Exit Sub
        End If
        
        Order_no_search.show
        Order_no_search.RetrunType = 10
Order_no_search.Label1(2).Caption = transactionName
                 Order_no_search.lblSpecificsearch = transactiontype
                        
                        If val(Me.DBCboClientName.BoundText) <> 2 Then
                        
                            Order_no_search.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
                        End If
    
    
    
    End If


End Sub

Private Sub txtempcode_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim EmpID As Integer

    If KeyCode = vbKeyReturn Then
        GetEmployeeIDFromCode txtempcode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If


End Sub

Private Sub TxtPONo_Change()
  If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        RetriveOrder Me.TxtPONo
    End If
End Sub

Private Sub TxtPONo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Order_no_search4.show
        Order_no_search4.RetrunType = 43

        If val(Me.DBCboClientName.BoundText) <> 2 Then
        
            Order_no_search4.DBCboClientName.BoundText = Me.DBCboClientName.BoundText
        End If
    End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 2
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TxtLcNo_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Order_no_search3.show
        Order_no_search3.RetrunType = 1
         
    End If
        
End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId
    End If
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
'    On Error GoTo ErrTrap

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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            '        Cmd_Click (0)
        Else
            '        SendKeys "{TAB}"
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
        '    If Cmd(3).Enabled = False Then Exit Sub
        '    Cmd_Click (3)
    End If

    If KeyCode = vbKeyF8 Then
        If Cmd(4).Enabled = False Then Exit Sub
        Cmd_Click (4)
    End If

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
        
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
       
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.Text = "N" Or TxtModFlg.Text = "E" Then
            
            End If
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_Change()
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset
    Dim ShowTax As Boolean
    Dim Dcombos As ClsDataCombos
   Dim My_SQL2 As String
'    On Error GoTo ErrTrap

   ' If GeneralPriceType = 0 Then
        ScreenNameArabic = "  ÿ·»«  œ«Œ·Ì… "
        ScreenNameEnglish = "Internal Order "
        CurrentTransactionType = 38
  
   ' End If
 My_SQL2 = " select id,Project_name from projects order by Project_name"
    fill_combo DCPROJECT1, My_SQL2

    My_SQL2 = " select  oprid,des from projects_des"
    fill_combo Dcterm1, My_SQL2

    My_SQL2 = " select  id,name from terms_operations"
    fill_combo dcopr, My_SQL2
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Me.Caption = ScreenNameArabic
    Ele(6).Caption = ScreenNameArabic

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang

    End If
    
    StrSQL = "SELECT * From groups "
    fill_combo XPCboGroupBuiltin, StrSQL



    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    Ele(4).Visible = ShowTax
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set CmdConvert.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    Set CmdTemplate.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
     Set Accredit.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Required").Picture
    Set NewGrid.Grid = FG
    NewGrid.GridTrans = GridTransType.internalorder
    
    
        Dim intDefStore  As Integer
     intDefStore = 0
GetUserData user_id, , , , , , , , , intDefStore


    
    
    
   
    NewGrid.DefStorePurchase = intDefStore
    
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    'Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
  Set NewGrid.DtpBillDate = Me.XPDtbBill
  
    ' ⁄»∆… »Ì«‰«  «·√’‰«›

    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
        Set NewGrid.StoreName = DCboStoreName
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    'Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    ' Resize_Form Me, TransactionSize
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    FG.WallPaper = BGround.Picture
    AddTip
    XPDtbBill.value = Date
    Set Dcombos = New ClsDataCombos

     Dcombos.GetEmployees Me.DcboEmpName
        Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True '  2 supplier  1 customer
        Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName1, True '  2 supplier  1 customer
        
     Dcombos.GetEmpDepartments Me.DcboEmpDepartments
     Dcombos.GetEquipments DCEquipments


    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetStores Me.DCboStoreName1
    
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetBranches Me.dcBranch

    If SystemOptions.usertype <> UserAdminAll Then
            If checkmanyBranches = False Then
                   Me.dcBranch.Enabled = True
             End If
    
      If checkmanyStores = False Then
                   Me.DCboStoreName.Enabled = True
             End If
             
    End If
    
    
    
    '///////////
    With Combo1
    If SystemOptions.UserInterface = ArabicInterface Then
    .Clear
    .AddItem ("„ﬁ»Ê·")
    .AddItem ("„—›Ê÷")
    Else
     .Clear
    .AddItem ("Accepted")
    .AddItem ("Refused")
    End If
    End With
    
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName

    Dcombos.GetSalesRepData Me.DcboEmp
 
    Set cSearchDcbo(3) = New clsDCboSearch
    Set cSearchDcbo(3).Client = Me.DcboEmp
    cSearchDcbo(3).SetBuddyText Me.TxtEmployeeID

    NewGrid.FillGrid
With CboPriceType
.AddItem "  ⁄«œÌ "
End With

    With Me.CBOInternalFlag
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
             
            .AddItem "  ⁄«œÌ "
       .AddItem "  ÿ«—Ì¡ "
  
       .AddItem "   √ﬂÌœ "
      .AddItem "  «’·«Õ "
 
       
        Else
             
            .AddItem " Routine "
 .AddItem " Critical  "
 .AddItem " Import "
 .AddItem " Confirmation "
 .AddItem " Repair "
 .AddItem " Local "
 
        End If

        .ListIndex = 0
    End With



    With Me.purchaseType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
             
            .AddItem "  „Õ·Ì "
 .AddItem "  «” Ì—«œ "
      
       
        Else
             
            .AddItem " Local "
            .AddItem " Import "
 
        End If

        .ListIndex = 0
    End With



'purchaseType

    With Me.CboType
        .Clear

               If SystemOptions.UserInterface = ArabicInterface Then
                   .AddItem "   ÌœÊÌ "
                   .AddItem "«·Ì ÿ»ﬁ« ·Õœ «·ÿ·» "
            
               Else
                   .AddItem "Manual"
                   .AddItem "Auto "
            
               End If

        .ListIndex = 0
    End With
    
    If SystemOptions.UserInterface = ArabicInterface Then
    
     With Me.DCOPrType
        .Clear
        .AddItem "»·«"
        .AddItem "„Ê«œ Œ«„ "
        .AddItem "„Â„« "
        .AddItem "ﬁÿ⁄ €Ì«— "
        .AddItem "«‰ «Ã  «„ "
        .AddItem "Âœ«Ì« Ê⁄Ì‰« "
        .AddItem "⁄Âœ…/ «’Ê· À«» …  "
    End With
    
    With CBOOrderType
    .Clear
        .AddItem "’—›"
        .AddItem " ÕÊÌ·"
        .AddItem "«—Ã«⁄ ·„Ê—œ"
        .AddItem "‘—«¡"
        .AddItem "«—Ã«⁄ œ«Œ·Ì"
        
    End With
    
    With CBoBasedON
    .Clear
        .AddItem "»·«"
        .AddItem "  «„— »Ì⁄"
        .AddItem " Œÿ… „»Ì⁄«  "
        .AddItem " Œÿ… «‰ «Ã "
    .AddItem " √„— ‘€·   "
    .AddItem " ≈—Ã«⁄ œ«Œ·Ì"
        .AddItem " ÿ·» œ«Œ·Ì"

   End With
    Else
    
       With Me.DCOPrType
        .Clear
        .AddItem "NA"
        .AddItem "RM "
        .AddItem "Missions"
        .AddItem "Spare Part"
        .AddItem "F.P."
        .AddItem "Gifts"
        .AddItem "F.A"
    End With
    
    With CBOOrderType
    .Clear
        .AddItem "Issue"
        .AddItem "Transfer"
        .AddItem "Vendor Return"
        .AddItem "Purchase"
        .AddItem "«Internal Return"
        
    End With
    
    With CBoBasedON
    .Clear
        .AddItem "Na"
        .AddItem "PO"
        .AddItem "Sales Plan"
        .AddItem "Production Plan"
    .AddItem "Job Order"
    .AddItem "Internal Return"
        .AddItem " Internal Request"

   End With
   
   End If
    'StrSQL = "SELECT * FROM Transactions WHERE (Transaction_Type=6 or Transaction_Type=29  or Transaction_Type=17)" 'OR Transaction_Type=17
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=" & CurrentTransactionType
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"

       If SystemOptions.usertype <> UserAdmin Then
          '      StrSQL = StrSQL & " AND   BranchId=" & Current_branch
            End If
            
    StrSQL = StrSQL + " Order By Transaction_ID"
        
        
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim My_SQL As String
    My_SQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, My_SQL
    'fill_combo Me.DcCostCenter, My_SQL

    My_SQL = " select code,account_name from markaas_taklefa"
 
    fill_combo Me.DcCostCenter, My_SQL

    My_SQL = " select id,Project_name from projects"
 
    fill_combo Me.DataCombo2, My_SQL

    My_SQL = " select CountryID,CountryName from TblCountriesData"
 
    fill_combo Me.DataCombo4, My_SQL

    My_SQL = " select id,name from Shipment_mode"
 
    fill_combo Me.DataCombo5, My_SQL
    CboPriceType.ListIndex = GeneralPriceType

    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If
    FG.Editable = flexEDKbdMouse
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set SaleReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·”‰œ   " & TXT_order_no.Text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & "«‰Ê⁄ «·”‰œ  " & CboPriceType.Text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.Text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.Text & CHR(13) & " —ﬁ„ «·«⁄ „«œ    " & TxtLcNo
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Vchr . No   " & TXT_order_no.Text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Type  " & CboPriceType.Text & CHR(13) & " Store  " & DCboStoreName.Text & CHR(13) & " Customer/ Supplier " & DBCboClientName.Text & CHR(13) & " Lc NO    " & TxtLcNo
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , , Me.TXT_order_no
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , , Me.TXT_order_no
    End If
    
End Function

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"
            ' Me.Caption = "⁄—÷ √”⁄«—"
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPBtnNewClients.Enabled = False
          FG.Editable = flexEDKbdMouse
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            Me.DCboStoreName1.locked = True
        '    FG.Editable = flexEDNone
      '      Accredit.Enabled = True
            CmdConvert.Enabled = True
            '   CmdConvert.Visible = True
            CmdTemplate.Visible = False

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
                CmdConvert.Enabled = False
                Accredit.Enabled = False
            End If

            Ele(2).Enabled = False

If CBOOrderType.ListIndex = 1 Then
lbl(3).Visible = True
'TxtStoreID1.Visible = True
'DCboStoreName1.Visible = True
Else
lbl(3).Visible = False
'TxtStoreID1.Visible = False
'DCboStoreName1.Visible = False
End If

        Case "N"
            ' Me.Caption = "⁄—÷ √”⁄«—( ÃœÌœ )"
            Accredit.Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            Accredit.Enabled = False
            '   Me.XPBtnMove(0).Enabled = False
            '   Me.XPBtnMove(1).Enabled = False
            '   Me.XPBtnMove(2).Enabled = False
            '   Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.Rows = 2
            Me.XPDtbBill.Enabled = True
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.DCboStoreName1.locked = False
            FG.Editable = flexEDKbdMouse
        
            CmdConvert.Visible = False
            CmdTemplate.Enabled = True
            '  CmdTemplate.Visible = True
            Ele(2).Enabled = True
            CboItemCase.ListIndex = 0


        Case "E"
            ' Me.Caption = "⁄—÷ √”⁄«—(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            FG.Enabled = True
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.DCboStoreName1.locked = False
            FG.Editable = flexEDKbdMouse
            XPBtnNewClients.Enabled = True
        
            Accredit.Enabled = False
            CmdConvert.Visible = False
            CmdTemplate.Visible = False
            Ele(2).Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim Num As Long
    Dim Dusername As String
    On Error GoTo ErrTrap
TxtNoteSerial1.Text = ""
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.Find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If
'Me.TxtModFlg.text = "R"
    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    
    Combo1.ListIndex = IIf(ISNULL(rs("Shipping_Pos").value), -1, (rs("Shipping_Pos").value))
    
    XPTxtBillID.Text = IIf(ISNULL(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    Me.DcboEmpName.BoundText = IIf(ISNULL(rs("Emp_ID").value), "", rs("Emp_ID").value)
    TXT_order_no.Text = IIf(ISNULL(rs("order_no").value), "", rs("order_no").value)
    TxtPONo.Text = IIf(ISNULL(rs("PONo").value), "", rs("PONo").value)

    If rs("shipped").value = True Then
        chkshipped.value = vbChecked
    Else
        chkshipped.value = Unchecked
    End If

    Me.DataCombo4.BoundText = IIf(ISNULL(rs("countryid").value), "", rs("countryid").value)
        TxtBillComment.Text = IIf(ISNULL(rs("TransactionComment").value), "", (rs("TransactionComment").value))
    TxtTransSerial.Text = IIf(ISNULL(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(ISNULL(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    Me.DBCboClientName.BoundText = IIf(ISNULL(rs("CusID").value), "", rs("CusID").value)
    
    Me.DBCboClientName1.BoundText = IIf(ISNULL(rs("CusID1").value), "", rs("CusID1").value)
    
    Dccurrency.BoundText = IIf(ISNULL(rs("Currency_id").value), "", rs("Currency_id").value)
    DCEquipments.BoundText = IIf(ISNULL(rs("FixesAssetsID").value), "", rs("FixesAssetsID").value)
    DcboEmpDepartments.BoundText = IIf(ISNULL(rs("DepartementID").value), "", rs("DepartementID").value)
    
    'If rs("Transaction_Type").value = 6 Then
    '    Me.CboPriceType.ListIndex = 1
    'ElseIf rs("Transaction_Type").value = 17 Then '17
    '    Me.CboPriceType.ListIndex = 0
    'ElseIf rs("Transaction_Type").value = 29 Then
    'Me.CboPriceType.ListIndex = 2
    'End If

     If Not ISNULL(rs("general_cost_center").value) Then
        Me.DcCostCenter.BoundText = IIf(rs("general_cost_center").value = "", "", rs("general_cost_center").value)
    Else
        Me.DcCostCenter.BoundText = ""
    End If


 '  Me.CboPriceType.ListIndex = 0
    Me.DCboUserName.BoundText = IIf(ISNULL(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(ISNULL(rs("StoreID").value), "", rs("StoreID").value)
Me.DCboStoreName1.BoundText = IIf(ISNULL(rs("StoreID1").value), "", rs("StoreID1").value)
'55555555555555555555555555555555
    Me.TxtAddress.Text = IIf(ISNULL(rs("Address").value), "", (rs("Address").value))
 Me.TxtContactPhone.Text = IIf(ISNULL(rs("ContactPhone").value), "", (rs("ContactPhone").value))

            TxtBillComment.Text = IIf(ISNULL(rs("TransactionComment").value), "", (rs("TransactionComment").value))

    If Not (ISNULL(rs("CashCustomerPhone").value)) Then
        Me.TxtPhone.Text = rs("CashCustomerPhone").value
    Else
        Me.TxtPhone.Text = ""
    End If


    If Not (ISNULL(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If
Dim ContactTime As Date
 If Not ISNULL(rs("ContactTime").value) Then
        ContactTime = FormatDateTime(rs("ContactTime").value, vbShortTime)
        Me.DpContactTime.value = ContactTime
   
    End If

       

'55555555555555555555555555555555
    XPTxtTaxValue.Text = IIf(ISNULL(rs("TaxValue").value), "", (rs("TaxValue").value))
    TxtLcNo.Text = IIf(ISNULL(rs("LcNo").value), "", (rs("LcNo").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)
    dcBranch.BoundText = IIf(ISNULL(rs("BranchId").value), "", rs("BranchId").value)

CBOInternalFlag.ListIndex = IIf(ISNULL(rs("InternalFlag").value), 0, rs("InternalFlag").value)
purchaseType.ListIndex = IIf(ISNULL(rs("purchaseType").value), 0, rs("purchaseType").value)

'purchaseType

CBOOrderType.ListIndex = IIf(ISNULL(rs("OrderType").value), 0, rs("OrderType").value)
DCOPrType.ListIndex = IIf(ISNULL(rs("OPrType").value), 0, rs("OPrType").value)
CBoBasedON.ListIndex = IIf(ISNULL(rs("BillBasedOn").value), 0, rs("BillBasedOn").value)

 
    Me.TxtNoteSerial1.Text = IIf(ISNULL(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.Text = IIf(ISNULL(rs("OldNoteSerial1").value), IIf(ISNULL(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    If TXT_order_no <> "" Then
 '       Me.TxtNoteSerial1.Text = TXT_order_no
    End If

    'Txt_order_no

'    lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
   ''///12 05 2015
   Me.TxtPeriods.Text = IIf(ISNULL(rs("Priod").value), "", rs("Priod").value)
   Me.DcbPeriodsID.ListIndex = IIf(ISNULL(rs("PriodDMY").value), -1, rs("PriodDMY").value)

'    DBPix202.ImageClear

'    If Dir(App.path & "\images\sign\sign" & rs("posted").value & ".JPG") <> "" Then
'
'        DBPix202.ImageLoadFile (App.path & "\images\sign\sign" & user_id & ".JPG")
'    End If

   If ISNULL(rs("posted").value) Then
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "   «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " send to Approval   "
                                               End If
                                               Accredit.Enabled = True
  Else
                                                   If SystemOptions.UserInterface = ArabicInterface Then
                                                    Accredit.Caption = "  „ «·«—”«· ··«⁄ „«œ "
                                                  Else
                                                    Accredit.Caption = " sent to Approval   "
                                               End If
                                               Accredit.Enabled = False
   End If
   
    'If Not IsNull(rs("posted").value) Then
    '    Frame4.Visible = True
    '    GetUserData val(rs("posted").value), , , , , , , Dusername
    '    LblPostedPerson = Dusername

    '                If user_id = rs("posted").value Then
    '                                If CheckOrderNotInTransaction(21, TxtNoteSerial1) = False Then
    '                                                If SystemOptions.UserInterface = ArabicInterface Then
    '                                                    Accredit.Caption = "«·€«¡ «·«⁄ „«œ "
    '                                                Else
    '                                                    Accredit.Caption = "Cancel Accredit   "
    '                                                End If
    '
    '                                Else
    '
    '                                                If SystemOptions.UserInterface = ArabicInterface Then
    '                                                    Accredit.Caption = "  «—”«· ··«⁄ „«œ "
    '                                                Else
    '                                                    Accredit.Caption = " send to accredit   "
     '                                               End If
    '
    '                                End If
    '
    '                End If

    'Else
    '    Frame4.Visible = False
    '    Accredit.Caption = "     «—”«· ··«⁄ „«œ "
    'End If
  
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = " SELECT    dbo.Transaction_Details.UnitId as vunitid,   dbo.TblItems.HaveSerial AS Expr1,Groups.GroupName, *"
    StrSQL = StrSQL + "  FROM         dbo.TblItems INNER JOIN"
    StrSQL = StrSQL + "                   dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
    StrSQL = StrSQL + "                  dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblProcessDEF ON dbo.Transaction_Details.Oper_ID = dbo.TblProcessDEF.TblProcessDEFID LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects_des ON dbo.Transaction_Details.Pand_ID = dbo.projects_des.oprid LEFT OUTER JOIN"
     StrSQL = StrSQL + "                  dbo.projects ON dbo.Transaction_Details.project_ID1 = dbo.projects.id"
    StrSQL = StrSQL + "                  LEFT OUTER JOIN     Groups ON dbo.Transaction_Details.GroupID = dbo.Groups.GroupID "

    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
StrSQL = StrSQL + " order by Transaction_Details.id "
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
           ''//
           
           


         FG.TextMatrix(Num, FG.ColIndex("TransactionID4")) = IIf(ISNULL(RsDetails("TransactionID4")), "", (RsDetails("TransactionID4").value))
         FG.TextMatrix(Num, FG.ColIndex("NoteSerial14")) = IIf(ISNULL(RsDetails("NoteSerial14")), "", (RsDetails("NoteSerial14").value))

         FG.TextMatrix(Num, FG.ColIndex("catalog")) = IIf(ISNULL(RsDetails("catalog")), "", (RsDetails("catalog").value))
         FG.TextMatrix(Num, FG.ColIndex("GroupID")) = IIf(ISNULL(RsDetails("GroupID")), "", (RsDetails("GroupID").value))
         FG.TextMatrix(Num, FG.ColIndex("GroupName")) = IIf(ISNULL(RsDetails("GroupName")), "", (RsDetails("GroupName").value))
         FG.TextMatrix(Num, FG.ColIndex("projectid")) = IIf(ISNULL(RsDetails("project_ID1")), "", (RsDetails("project_ID1").value))
         FG.TextMatrix(Num, FG.ColIndex("pandid")) = IIf(ISNULL(RsDetails("Pand_ID")), "", (RsDetails("Pand_ID").value))
         FG.TextMatrix(Num, FG.ColIndex("proid")) = IIf(ISNULL(RsDetails("Oper_ID")), "", (RsDetails("Oper_ID").value))
         FG.TextMatrix(Num, FG.ColIndex("pand")) = IIf(ISNULL(RsDetails("des")), "", (RsDetails("des").value))
         FG.TextMatrix(Num, FG.ColIndex("ShowAttatch")) = IIf(ISNULL(RsDetails("ShowAttatch")), "", (RsDetails("ShowAttatch").value))
         
         If SystemOptions.UserInterface = ArabicInterface Then
         FG.TextMatrix(Num, FG.ColIndex("project")) = IIf(ISNULL(RsDetails("Project_name")), "", (RsDetails("Project_name").value))
         FG.TextMatrix(Num, FG.ColIndex("pro")) = IIf(ISNULL(RsDetails("ProcessName")), "", (RsDetails("ProcessName").value))
         Else
         FG.TextMatrix(Num, FG.ColIndex("project")) = IIf(ISNULL(RsDetails("Project_nameE")), "", (RsDetails("Project_nameE").value))
         FG.TextMatrix(Num, FG.ColIndex("pro")) = IIf(ISNULL(RsDetails("ProcessNameE")), "", (RsDetails("ProcessNameE").value))
         End If
        ''//
    
        
        FG.TextMatrix(Num, FG.ColIndex("RecivePeriod")) = IIf(ISNULL(RsDetails("RecivePeriod")), "", (RsDetails("RecivePeriod").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(ISNULL(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(ISNULL(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(ISNULL(RsDetails("showqty")), "", (RsDetails("showqty").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(ISNULL(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(ISNULL(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(ISNULL(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(ISNULL(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(ISNULL(RsDetails("vunitid")), "", (RsDetails("vunitid").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(ISNULL(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(ISNULL(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(ISNULL(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(ISNULL(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        




        FG.TextMatrix(Num, FG.ColIndex("RequestLimit")) = IIf(ISNULL(RsDetails("RequestLimit")), 0, (RsDetails("RequestLimit").value))
        FG.TextMatrix(Num, FG.ColIndex("LastPurchaseDate")) = IIf(ISNULL(RsDetails("LastPurchaseDate")), "", (RsDetails("LastPurchaseDate").value))
        FG.TextMatrix(Num, FG.ColIndex("LastPurchasePrice")) = IIf(ISNULL(RsDetails("LastPurchasePrice")), 0, (RsDetails("LastPurchasePrice").value))
        FG.TextMatrix(Num, FG.ColIndex("LastPurchaseqty")) = IIf(ISNULL(RsDetails("LastPurchaseqty")), 0, (RsDetails("LastPurchaseqty").value))
        FG.TextMatrix(Num, FG.ColIndex("AverageIssue")) = IIf(ISNULL(RsDetails("AverageIssue")), 0, (RsDetails("AverageIssue").value))
        FG.TextMatrix(Num, FG.ColIndex("AverageIssueyraly")) = IIf(ISNULL(RsDetails("AverageIssueyraly")), 0, (RsDetails("AverageIssueyraly").value))
        
        ''//12 05 2015
         FG.TextMatrix(Num, FG.ColIndex("ItemBalance")) = IIf(ISNULL(RsDetails("ItemBalance")), 0, (RsDetails("ItemBalance").value))
            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If
fillapprovData
    TxtFillData.Text = "F"
    FG.Editable = flexEDKbdMouse
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
            cmdCreateProduction.Enabled = True
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
Function fillapprovData()
      Dim Num As Integer
       Dim RsDetails As New ADODB.Recordset
       Dim StrSQL As String
      '    StrSQL = "SELECT     TOP 100 PERCENT  dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, dbo.ApprovalData.currorder, "
      'StrSQL = StrSQL + "  dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks, dbo.TblEmployee.Emp_Code,"
      'StrSQL = StrSQL + "   dbo.TblEmployee.emp_name , dbo.TblEmployee.Emp_Namee, dbo.TbLLevels.name, dbo.TbLLevels.namee"
      'StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
      'StrSQL = StrSQL + "   dbo.TblEmployee ON dbo.ApprovalData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
      'StrSQL = StrSQL + "   dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID"
      'StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(XPTxtBillID.text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.name & "')"
      'StrSQL = StrSQL + "  ORDER BY dbo.ApprovalData.levelorder"
       
10     StrSQL = "SELECT     TOP 100 PERCENT dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
20    StrSQL = StrSQL + " dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
30    StrSQL = StrSQL + " dbo.TbLLevels.name , dbo.TbLLevels.namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName"
40    StrSQL = StrSQL + " FROM         dbo.ApprovalData INNER JOIN"
50    StrSQL = StrSQL + " dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
60    StrSQL = StrSQL + " dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID"
70    StrSQL = StrSQL + " WHERE     (dbo.ApprovalData.Transaction_ID = " & val(XPTxtBillID.Text) & ") AND (dbo.ApprovalData.ScreenName = N'" & Me.Name & "')"
80    StrSQL = StrSQL + " ORDER BY dbo.ApprovalData.levelorder"

90        RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

100    If Not (RsDetails.EOF Or RsDetails.BOF) Then
110           GRID2.Rows = RsDetails.RecordCount + 1
       

120           For Num = 1 To RsDetails.RecordCount
              
130          GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(ISNULL(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
140       If GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = "1" Then
150      GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = &HFFFFC0
160      Else
170       GRID2.Cell(flexcpBackColor, Num, 1, Num, 7) = vbWhite
180       End If
          
190           GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(ISNULL(RsDetails("ApprovDate")), "", flexChecked)
200              If SystemOptions.UserInterface = ArabicInterface Then
210               GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(ISNULL(RsDetails("Name")), "", Trim(RsDetails("Name").value))
220             Else
230                GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(ISNULL(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
240             End If
250               If SystemOptions.UserInterface = ArabicInterface Then
260               GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(ISNULL(RsDetails("UserName")), "", (RsDetails("UserName").value))
270               Else
280               GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(ISNULL(RsDetails("UserName")), "", (RsDetails("UserName").value))
290               End If
300               GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(ISNULL(RsDetails("ApprovDate")), "", (RsDetails("ApprovDate").value))
310             GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(ISNULL(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
       
       
320   RsDetails.MoveNext
330   If Num = RsDetails.RecordCount Then

340           If GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) <> "" Then
350                                   If SystemOptions.UserInterface = ArabicInterface Then
360                                         Label11.Caption = " „ «·«⁄ „«œ ··„” ‰œ »«·ﬂ«„·"
370                                    Else
380                                          Label11.Caption = "Approved"
390                                    End If
400                               Label11.backcolor = &H80FF80
410           Else
420                                If SystemOptions.UserInterface = ArabicInterface Then
430                                        Label11.Caption = "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
440                               Else
450                                        Label11.Caption = "Currently required Approve"
460                               End If
470                    Label11.backcolor = &HFFFFC0
480           End If

490   End If

500           Next Num
510   Else
520    GRID2.Rows = 1
530       End If
540   RsDetails.Close

End Function

Private Sub XPCboGroupBuiltin_Click(Area As Integer)
    Dim Dcombos As New ClsDataCombos
    Dim mIndex As Integer
    If Trim(XPCboGroupBuiltin.BoundText) <> "" Then
        mIndex = val(XPCboGroupBuiltin.BoundText)
        Dcombos.GetItemsNamesupdate Me.DCboItemsName, , , , , mIndex
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2, , , , , mIndex
    Else
        Dcombos.GetItemsNamesupdate Me.DCboItemsName
        'Dcombos.GetItemsNamesupdate Me.DcboItemID2
    End If

End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap
 
    Me.LblTotal.Caption = XPTxtSum.Text
 
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.Text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"
            Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.Find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.Text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If XPTxtBillID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (XPTxtBillID.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
Dim StrSQL As String

                  StrSQL = "Delete Transactions Where Transaction_ID In (Select TransactionID4 From Transaction_Details DD Where Transaction_ID = " & val(rs("Transaction_ID").value) & " )"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                    StrSQL = "Delete Transaction_Details Where Transaction_ID In (Select Transaction_Details Where Transaction_ID = " & val(rs("Transaction_ID").value) & " )"
                    Cn.Execute StrSQL, , adExecuteNoRecords
                '
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«    ÿ·» œ«Œ·Ì   ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷  ﬁ—Ì— »«·»Ì«‰«  «·Õ«·Ì… " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«   ÿ·» œ«Œ·Ì   «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄—÷ «·”⁄— «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«   ÿ·» œ«Œ·Ì   «·Õ«·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄—÷ ”⁄—" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "⁄—÷ √”⁄«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub SaveData()
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
    'Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim BeginTrans As Boolean
  On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.Text <> "R" Then
        If DBCboClientName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
'                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ê—œ"
            Else
'                Msg = "Please Select Vendor"
            End If

'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            DBCboClientName.SetFocus
'            SendKeys "{F4}"
            'Screen.MousePointer = vbDefault
            'Exit Sub
        End If

        If DCboStoreName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            Else
                Msg = "Select Inventory"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DCboStoreName.SetFocus
           SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        If Dccurrency.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "Õœœ «·⁄„·…"
            Else
                Msg = "Select Currency"
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Dccurrency.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
'        If Me.CboPriceType.ListIndex = -1 Then
     '       If SystemOptions.UserInterface = ArabicInterface Then
     '           Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄    «·«„—  ( )...!!!"
     '       Else
     '           Msg = "Specify Order Type"
     '       End If
'
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            CboPriceType.SetFocus
'            SendKeys "{F4}"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If

        If XPChkTAX.value = Checked Then
            If XPTxtTaxValue.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "ÌÃ» «œŒ«· ﬁÌ„… ÷—Ì»… «·„»Ì⁄« "
                Else
                    Msg = "Insert Sales Tax"
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                XPTxtTaxValue.SetFocus
                FG.SetFocus
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    
 
    
        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If

        Set RSTransDetails = New ADODB.Recordset
      '  RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      
        Dim Transaction_Type As Integer
        Dim Sanad_No As Integer

        If Me.CboPriceType.ListIndex = 0 Then
            Transaction_Type = CurrentTransactionType
            Sanad_No = CurrentTransactionType
  
 
         
        End If

        my_branch = val(dcBranch.BoundText)

        If TxtNoteSerial1.Text = "" Then
            If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 0, , Transaction_Type, , val(DCboStoreName.BoundText)) = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ    " & CHR(13) & " Enter Vchr No": Exit Sub
                Else
                    TxtNoteSerial1.Text = Voucher_coding(val(my_branch), XPDtbBill.value, Sanad_No, 170, , Transaction_Type, , val(DCboStoreName.BoundText))
                End If
            End If
        End If
 
        'TXT_order_no = Me.TxtNoteSerial1.Text
 
        Cn.BeginTrans
        BeginTrans = True
    
        If Me.TxtModFlg.Text = "N" Then
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=6"))
            
            rs.AddNew
        End If

        Screen.MousePointer = vbArrowHourglass
        
        rs("Shipping_Pos") = IIf(Combo1.ListIndex = -1, Null, Combo1.ListIndex)
        
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rs("branchID").value = val(Me.dcBranch.BoundText)
   
        rs("Transaction_ID").value = val(XPTxtBillID.Text)
        rs("order_no").value = TXT_order_no.Text
    
    If val(CBoBasedON.ListIndex) = 4 Then
    rs("OldOpOrderID").value = val(TXT_order_no.Text)
    Else
    rs("OldOpOrderID").value = Null
    End If
    
    
        If chkshipped.value = vbChecked Then
            rs("shipped").value = 1
        Else
            rs("shipped").value = 0
        End If
    
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Serial").value = TxtTransSerial.Text

      rs("PONO").value = IIf(TxtPONo.Text = "", Null, (TxtPONo.Text))
rs("Transaction_Type").value = CurrentTransactionType

rs("InternalFlag").value = val(CBOInternalFlag.ListIndex)
rs("purchaseType").value = val(purchaseType.ListIndex)


rs("BillBasedOn").value = val(CBoBasedON.ListIndex)

rs("OPrType").value = val(DCOPrType.ListIndex)
rs("OrderType").value = val(CBOOrderType.ListIndex)
  rs("ContactTime").value = FormatDateTime(Me.DpContactTime.value, vbShortTime)



    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))
        rs("UserID").value = user_id
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("CusID1").value = IIf(DBCboClientName1.BoundText = "", Null, val(DBCboClientName1.BoundText))
        
        rs("countryid").value = IIf(DataCombo4.BoundText = "", Null, val(DataCombo4.BoundText))
    
        rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
    
    rs("FixesAssetsID").value = IIf(DCEquipments.BoundText = "", Null, val(DCEquipments.BoundText))
    rs("DepartementID").value = IIf(DcboEmpDepartments.BoundText = "", Null, val(DcboEmpDepartments.BoundText))
    
 

        rs("Emp_ID").value = IIf(DcboEmpName.BoundText = "", Null, DcboEmpName.BoundText)
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("StoreID1").value = IIf(DCboStoreName1.BoundText = "", Null, val(DCboStoreName1.BoundText))
        
        rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
        rs("TaxValue").value = IIf(XPTxtTaxValue.Text = "", Null, val(XPTxtTaxValue.Text))
        rs("total").value = IIf(XPTxtSum.Text = "", Null, val(XPTxtSum.Text))
        rs("LcNo").value = IIf(TxtLcNo.Text = "", Null, (TxtLcNo.Text))
    
       rs("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
''///12 05 2015
rs("PriodDMY").value = val(Me.DcbPeriodsID.ListIndex)
rs("Priod").value = val(Me.TxtPeriods.Text)
'nnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnnn

'save
If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
    Else
        rs("CashCustomerName").value = Null
    End If

    If Trim$(Me.TxtPhone.Text) <> "" Then
        rs("CashCustomerPhone").value = Trim$(Me.TxtPhone.Text)
    Else
        rs("CashCustomerPhone").value = Null
    End If
    
    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))

    rs("ContactTime").value = FormatDateTime(Me.DpContactTime.value, vbShortTime)
       
            rs("Address").value = TxtAddress.Text
             rs("ContactPhone").value = TxtContactPhone.Text
 rs("Address").value = TxtAddress.Text


'nnnnnnnnnnnnnnnnnnnnnnn
        rs.update
    
        CuurentLogdata
  
        If Me.TxtModFlg.Text = "E" Then
        
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            
        End If

      StrSQL = "Delete Transactions Where Transaction_ID In (Select TransactionID4 From Transaction_Details DD Where Transaction_ID = " & val(rs("Transaction_ID").value) & " )"
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete Transaction_Details Where Transaction_ID In (Select Transaction_ID Transaction_Details Where Transaction_ID = " & val(rs("Transaction_ID").value) & " )"
        Cn.Execute StrSQL, , adExecuteNoRecords
    '
        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
                RSTransDetails("order_id").value = val(XPTxtBillID.Text)
         
                RSTransDetails("order_no").value = TXT_order_no.Text
             ''//
                


              RSTransDetails("TransactionID4").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TransactionID4")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("TransactionID4"))))
              RSTransDetails("NoteSerial14").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NoteSerial14")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("NoteSerial14"))))
            RSTransDetails("ShowAttatch").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ShowAttatch")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ShowAttatch"))))
                       
              RSTransDetails("GroupID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("GroupID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("GroupID"))))
              RSTransDetails("project_id1").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("projectid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("projectid"))))
              RSTransDetails("Pand_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("pandid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("pandid"))))
              RSTransDetails("Oper_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("proid")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("proid"))))
              RSTransDetails("catalog").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("catalog")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("catalog"))))
             ''/
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
               'RecivePeriod
               RSTransDetails("RecivePeriod").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("RecivePeriod")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("RecivePeriod"))))
               
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                RSTransDetails("ShowPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            
               RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), "", Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                 RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
                RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                
                RSTransDetails("RequestLimit").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("RequestLimit")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("RequestLimit"))))
              RSTransDetails("LastPurchaseDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LastPurchaseDate")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("LastPurchaseDate"))))
                RSTransDetails("LastPurchasePrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LastPurchasePrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("LastPurchasePrice"))))
                RSTransDetails("LastPurchaseqty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LastPurchaseqty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("LastPurchaseqty"))))
                RSTransDetails("AverageIssue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("AverageIssue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("AverageIssue"))))
                RSTransDetails("AverageIssueyraly").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("AverageIssueyraly")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("AverageIssueyraly"))))
                
                
               ''//12 05 2015
               RSTransDetails("ItemBalance").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemBalance")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemBalance"))))
                
                
 
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
                DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                  RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    'RSTransDetails("Price").value = Val(IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("Price")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("Price"))))) / RSTransDetails("Quantity").value
                    RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))) / RSTransDetails("QtyBySmalltUnit").value
                End If
                CalcCostPercent
                RSTransDetails("PercentCost").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("PercentCost")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("PercentCost")))))
                RSTransDetails.update
            End If

        Next RowNum

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
 'lbl(64).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
        Select Case Me.TxtModFlg.Text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & CHR(13)
                    Msg = Msg + "do you new Operation?"
        
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            
            Case "E"

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If
   cmdCreateProduction.Enabled = True
            
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
    
            Msg = "Cant Save Error"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    Else
        Msg = "Sorry... Error During Saving " & CHR(13)
    End If

    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub XPBtnNewClients_Click()

    'With FrmAddNewCustemer
    '    .DealingForm = ShowPrice
    '    .show vbModal
    '    .Caption = "≈÷«›… ⁄„Ì· ÃœÌœ"
    '    .lbl(1).Caption = "ﬂÊœ «·⁄„Ì·"
    '    .lbl(0).Caption = "«”„ «·⁄„Ì·"
    'End With

End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        XPTxtTaxValue.locked = False
        lbl(4).Enabled = True
    Else
        XPTxtTaxValue.Text = ""
        XPTxtTaxValue.Enabled = False
        lbl(4).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport()
    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice XPTxtBillID.Text, 8, DcboEmpName.Text
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim StrMSG As String
    Dim IntResult As String
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

Public Sub Cala()
    NewGrid.Calculate 1
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial1.Text = ""
 
End Sub

Private Sub XPTxtTaxValue_Change()

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
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
    chkshipped.Caption = "shipped"
    Label13.Caption = "Branch"
    lbl(56).Caption = "Opr Type"
     lbl(38).Caption = "Equipments"
          lbl(39).Caption = "Employee"
Label11.Caption = "Need Approval"
     lbl(47).Caption = "Status"
     lbl(48).Caption = "Project"
     lbl(43).Caption = "Item"
     lbl(51).Caption = "Process"
     lbl(45).Caption = "Client"
     Label14.Caption = "Cash Client"
     Label17.Caption = "Tel. No."
     Label16.Caption = "Call Time"
     lbl(44).Caption = "Address"
     Label15.Caption = "Tel."
     lbl(46).Caption = "from store"
     lbl(9).Caption = "Order Status"
     lbl(42).Caption = "Pur type"
     Label1(11).Caption = "inteval"
     ISButton2.Caption = "Attach."
     With DcbPeriodsID
     .Clear
     .AddItem "Day"
     .AddItem "Month"
     .AddItem "Year"
     End With
     
    Me.Caption = ScreenNameEnglish
    Ele(6).Caption = ScreenNameEnglish
    'Me.Caption = "Order Request/Proforma   Invoice"
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Approve Status"
    lbl(18).Caption = "Type"
    Label4.Caption = "ACC. BY"
    Label10.Caption = "Signature"
    lbl(32).Caption = "Sales Person"
    Accredit.Caption = "Accredit"
    Cmd(8).Caption = "Print Pur. Order"
    'Ele(6).Caption = Me.Caption
    lbl(50).Caption = "Discounts"
    lbl(49).Caption = "Net"

    lbl(5).Caption = "Ord/P INV. No"
    Frame3.Caption = "LC Data"
    ISButton1.Caption = "View"
    lbl(25).Caption = "Total"
    lbl(63).Caption = "Qty"
    Label2.Caption = "Branch"
    lbl(6).Caption = "Date"
    lbl(7).Caption = "Client"
    lbl(8).Caption = "Store"
    lbl(9).Caption = "Type"
    lbl(10).Caption = "Cost Center"
    lbl(11).Caption = "Project"
    lbl(16).Caption = "Article Section"
    lbl(12).Caption = "Currency"
    lbl(13).Caption = "Country"
    lbl(14).Caption = "Shipment Mode"
    lbl(17).Caption = "Value"
    lbl(15).Caption = "Payment M"
    lbl(28).Caption = "  Remarks"
    lbl(19).Caption = "Kind Of Order"
    lbl(24).Caption = "Expiry Date"
    lbl(20).Caption = "Credit Bank"
    lbl(21).Caption = "Credit Curr."
    lbl(22).Caption = "Credit No."
    lbl(23).Caption = "Value"
    'ISButton1.Caption = "Show Port Data"
    'Label1.Caption = "LC NO:"
    'Label2.Caption = "Supp info No."
    Label3.Caption = "Supp info Date"
    Label5.Caption = "Exp Del Date"
    Label6.Caption = "Act Del Date"
    Label7.Caption = "Late Date"
    Label8.Caption = "Exp Arrival Date"
    Label9.Caption = "Comments"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "item name"

    lbl(29).Caption = "Status"
    lbl(27).Caption = "Qty"
    lbl(26).Caption = "Price"

    lbl(3).Caption = "Total"
    lbl(1).Caption = "By"
    lbl(0).Caption = "Currenr rec."
    lbl(2).Caption = "Total rec."

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    CmdConvert.Caption = "Convert To Bill"
    CmdTemplate.Caption = "Insert template"







    With Me.GRID2
 
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "leve lName"

        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "ApprovDate"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"

    End With
 
    With Me.FG
 
        .TextMatrix(0, .ColIndex("ItemBalance")) = "Item Balance"
                 .TextMatrix(0, .ColIndex("RecivePeriod")) = "IRecive Period"
                         .TextMatrix(0, .ColIndex("project")) = "Project"
                                 .TextMatrix(0, .ColIndex("pand")) = "Term"
                                         .TextMatrix(0, .ColIndex("pro")) = "Process"
                                                 .TextMatrix(0, .ColIndex("ItemBalance")) = "Item Balance"
         .TextMatrix(0, .ColIndex("LastPurchasePrice")) = "Last Purchase Price"
         .TextMatrix(0, .ColIndex("LastPurchaseqty")) = "Las tPurchase qty"
         .TextMatrix(0, .ColIndex("RequestLimit")) = "Request Limit"
         .TextMatrix(0, .ColIndex("LastPurchaseDate")) = "Last Purchase Date"
         .TextMatrix(0, .ColIndex("AverageIssueyraly")) = "Yearly Average Issue "
         .TextMatrix(0, .ColIndex("AverageIssue")) = "Monthly  Average Issue"
         .TextMatrix(0, .ColIndex("FoxyNo")) = "Program No"
         
         
         
         
         
                                                 
                                                 

    End With
    
 
lbl(34).Caption = "Ordr Type:"
lbl(35).Caption = "Based On"
lbl(36).Caption = "Opr Type"
lbl(37).Caption = "Dept."
lbl(10).Caption = "Cost Center"

lbl(9).Caption = "Status"


End Sub


Private Sub CreateProduction(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreId As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String, ByVal mmID As Long, Transaction_ID As Long)

Dim BolTemp As Boolean
Dim sql As String
Dim Msg As String
Dim NoteID As Long

Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim StrSQL As String
Dim Percetage As Double
Dim AccountVATCreit As String
Dim mPrice As Double
Dim rsDummy As New ADODB.Recordset
' «·”⁄— Â‰« ÂÊ ’«›Ï «·”⁄— »⁄œ Œ’„ «·«÷«›Ï Ê«·Œ’Ê„« 
'
'PercentgValueAddedAccount_Transec XPDtbBill.value, 21, 0, AccountVATCreit, Percetage
'PercetageVat = Percetage

'BillTOTAL = 0




 
Dim RSTransDetails As New ADODB.Recordset
     
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
StrSQL = "Select Item_ID,UnitID,Sum(ShowQty) Qty,Sum(ShowPrice) Price,Sum(PercentCost) PercentCost from Transaction_Details Where ID = " & mmID
StrSQL = StrSQL & " Group By Item_ID,UnitID "
rsDummy.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic
    If Not rsDummy.EOF Then
    
        Dim mItemNo As Long, mUnitNo As Long, mQty As Long, mVAt2 As Double, mTotal As Double
        Dim mwidtj As Double, mhight As Double, mTotalAdd As Double, mTotalDisc As Double, mNet As Double, mTotalWithVat As Double, mLength As Double
        Dim mItemName2 As String, mCostPercent As Double
        Dim mRemark As String
        mItemNo = val(rsDummy!Item_ID & "")
        If mItemNo = 0 Then GoTo NextRow
        
               
            mItemNo = val(rsDummy!Item_ID & "")
           
            mUnitNo = val(rsDummy!UnitID & "")
            mQty = val(rsDummy!Qty & "")
            mPrice = val(rsDummy!Price & "")
'            mwidtj = val(rsDummy!widtj & "")
'            mhight = val(rsDummy!hight & "")
'            mLength = val(rsDummy!Length & "")
           ' mTotal = val(rsDummy!Total & "")
        '    mRemark = Trim(rsDummy!Remark & "")
        '    mTotalDisc = val(rsDummy!TotalDisc & "")
        '    mTotalAdd = val(rsDummy!TotalAdd & "")
        '    mNet = val(rsDummy!net & "")
        '    mVAt2 = val(rsDummy!Vat2 & "")
           ' mTotalWithVat = val(rsDummy!TotalWithVat & "")
          '  mPrice = (val(mTotal) + val(mTotalAdd)) / val(mQty)
            mCostPercent = val(rsDummy!PercentCost & "")
            
        RSTransDetails.AddNew
        RSTransDetails("Transaction_ID").value = Transaction_ID
        RSTransDetails("ColorID").value = 1
        RSTransDetails("ItemSize").value = 1
        RSTransDetails("ClassId").value = 1
        RSTransDetails("Item_ID").value = mItemNo
        RSTransDetails("UnitID").value = mUnitNo
        RSTransDetails("SHOWQTY").value = mQty
        RSTransDetails("PercentCost").value = mCostPercent
        RSTransDetails("showPrice").value = mPrice
        RSTransDetails("Lineexpenses").value = mPrice
        
        RSTransDetails("ItemDiscountType").value = 2
        
        If SystemOptions.TypicalProduction = False Then

            RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value, StoreId)

            If RSTransDetails("CostPrice").value = 0 Then
                RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , LastPurPriceType, , , XPDtbBill.value, val(Me.Text1.Text), RSTransDetails("UnitID").value, val(Me.DCboStoreName.BoundText))
                
            End If
              RSTransDetails("CostPrice").value = mPrice
        Else
            RSTransDetails("CostPrice").value = 0
        
        End If
                      
          
                      '«·ÊÕœ« 
       
        Dim RsUnitData As ADODB.Recordset
        Dim LngCurItemID As Long
        Dim LngUnitID As Long
        Dim DblQty As Double
    
        LngCurItemID = val(mItemNo)
        LngUnitID = val(mUnitNo)
        DblQty = val(mQty)

        StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
        StrSQL = StrSQL + " AND UnitID=" & LngUnitID
        Set RsUnitData = New ADODB.Recordset
        RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
            RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
            RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
            RSTransDetails("OpeningSalesValue").value = RSTransDetails("CostPrice").value * val(mQty)
            RSTransDetails("Price").value = val(IIf((mPrice = 0), 0, val(mPrice))) / RSTransDetails("QtyBySmalltUnit").value
        
        End If

    
         UpdateTransactionsCost CStr(Transaction_ID)
         RSTransDetails.update
    
      '  Dim i As Integer
        'Dim sql As String
    End If
NextRow:


NoteSerial = Notes_coding(val(BranchID), Transaction_Date)






'***********************
'End If
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:



 

End Sub



 Private Sub CalcCostPercent()
    Dim i As Long
    Dim mCostPercent As Double
    Dim mCostTotal As Double
    mCostTotal = FG.Aggregate(flexSTSum, FG.FixedRows, FG.ColIndex("Price"), FG.Rows - 1, FG.ColIndex("Price"))
    For i = 1 To FG.Rows - 1
    If mCostTotal <> 0 Then
        FG.TextMatrix(i, FG.ColIndex("PercentCost")) = val(FG.TextMatrix(i, FG.ColIndex("Price"))) / mCostTotal * 100
        Else
        FG.TextMatrix(i, FG.ColIndex("PercentCost")) = 0
        End If
    Next
    
 End Sub


