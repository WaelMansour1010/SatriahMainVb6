VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReturnSalling33 
   Caption         =   "„—œÊœ« «·„»Ì⁄« "
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14400
   HelpContextID   =   360
   Icon            =   "FrmReturnSalling33.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   14400
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
      Height          =   9165
      Left            =   0
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   0
      Width           =   14400
      _cx             =   25400
      _cy             =   16166
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
         Height          =   2325
         Index           =   0
         Left            =   15
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   645
         Width           =   14370
         _cx             =   25347
         _cy             =   4101
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
         Begin VB.TextBox TxtBillComment 
            Alignment       =   1  'Right Justify
            Height          =   450
            Left            =   5640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   163
            Top             =   1200
            Width           =   2340
         End
         Begin VB.TextBox TxtVATNO 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   6600
            TabIndex        =   160
            Top             =   825
            Width           =   1485
         End
         Begin VB.TextBox txt_Currency_rate 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Text            =   "1"
            Top             =   135
            Width           =   525
         End
         Begin VB.TextBox TxtManualNO 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   8475
            TabIndex        =   132
            Top             =   1560
            Width           =   1590
         End
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   8820
            TabIndex        =   4
            Top             =   1200
            Width           =   4365
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   11580
            TabIndex        =   108
            Top             =   1560
            Width           =   1605
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   1065
            TabIndex        =   104
            Top             =   480
            Width           =   915
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   2775
            Style           =   2  'Dropdown List
            TabIndex        =   103
            Top             =   480
            Width           =   1125
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   12240
            TabIndex        =   98
            Top             =   840
            Width           =   945
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12420
            TabIndex        =   97
            Top             =   1950
            Width           =   765
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  ﬁÌœ «·”‰œ"
            Height          =   645
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   1680
            Width           =   8430
            Begin VB.TextBox TxtValueTemp 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   120
               Visible         =   0   'False
               Width           =   1350
            End
            Begin VB.TextBox TxtManualNo1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   240
               Visible         =   0   'False
               Width           =   1350
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   5640
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   240
               Width           =   1305
            End
            Begin ImpulseButton.ISButton Cmd 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   10
               Left            =   4560
               TabIndex        =   95
               Top             =   240
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   450
               ButtonStyle     =   1
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
               ColorShadow     =   4210752
               ColorOutline    =   0
               DrawFocusRectangle=   0   'False
               ColorToggledHoverText=   16711680
               ColorTextShadow =   4210752
            End
            Begin ImpulseButton.ISButton Cmd 
               CausesValidation=   0   'False
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   96
               Top             =   240
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   450
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ ”‰œ «·«” ·«„"
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   285
               Index           =   36
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   240
               Width           =   2250
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
               Height          =   285
               Index           =   23
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   113
               Top             =   240
               Visible         =   0   'False
               Width           =   2250
            End
         End
         Begin VB.TextBox TxtEmployeeID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1335
            TabIndex        =   91
            Top             =   2955
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   80
            Top             =   1110
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1680
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   11370
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   120
            Width           =   1815
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   930
            Index           =   2
            Left            =   120
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   780
            Width           =   5520
            _cx             =   9737
            _cy             =   1640
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
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
            ForeColor       =   128
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "‰Ê⁄ ⁄„·Ì… «·√— Ã«⁄"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
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
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«” »œ«·"
               Height          =   225
               Index           =   1
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   180
               Width           =   1050
            End
            Begin VB.OptionButton opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„— Ã⁄"
               Height          =   225
               Index           =   0
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   180
               Width           =   1110
            End
            Begin VB.TextBox txtInvDate 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   1770
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   540
               Width           =   1095
            End
            Begin VB.TextBox TxtInvSerial 
               Alignment       =   1  'Right Justify
               Height          =   255
               Left            =   3585
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   480
               Width           =   960
            End
            Begin VB.TextBox TxtInvID 
               Alignment       =   1  'Right Justify
               Height          =   240
               Left            =   5700
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   390
               Visible         =   0   'False
               Width           =   450
            End
            Begin VB.ComboBox CboRetrunType 
               Height          =   315
               ItemData        =   "FrmReturnSalling33.frx":058A
               Left            =   2400
               List            =   "FrmReturnSalling33.frx":058C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   225
               Width           =   3015
            End
            Begin ImpulseButton.ISButton CmdSearchTrans 
               Height          =   255
               Left            =   360
               TabIndex        =   84
               Top             =   510
               Width           =   660
               _ExtentX        =   1164
               _ExtentY        =   450
               ButtonPositionImage=   1
               Caption         =   "»ÕÀ"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmReturnSalling33.frx":058E
            End
            Begin ImpulseButton.ISButton CmdOpenTrans 
               Height          =   255
               Left            =   1080
               TabIndex        =   66
               Top             =   510
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   450
               ButtonPositionImage=   1
               Caption         =   "⁄—÷"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmReturnSalling33.frx":0928
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌÕÂ«"
               Height          =   285
               Index           =   10
               Left            =   2895
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   540
               Width           =   600
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ ›« Ê—… «·»Ì⁄"
               Height          =   375
               Index           =   4
               Left            =   4500
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   480
               Width           =   900
            End
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   9540
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   -180
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   8730
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   450
            Width           =   1440
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   30
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   1485
            Visible         =   0   'False
            Width           =   2400
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   2625
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   1485
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   1125
            Visible         =   0   'False
            Width           =   990
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   8790
            TabIndex        =   3
            Top             =   840
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   8460
            TabIndex        =   15
            Top             =   1950
            Width           =   3930
            _ExtentX        =   6932
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   11295
            TabIndex        =   1
            Top             =   465
            Width           =   1890
            _ExtentX        =   3334
            _ExtentY        =   609
            _Version        =   393216
            Format          =   143785985
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   345
            Left            =   3765
            TabIndex        =   67
            Top             =   1155
            Visible         =   0   'False
            Width           =   720
            _ExtentX        =   1270
            _ExtentY        =   609
            ButtonPositionImage=   1
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmReturnSalling33.frx":0CC2
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   5160
            TabIndex        =   74
            Top             =   120
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   5160
            TabIndex        =   78
            Top             =   480
            Width           =   2925
            _ExtentX        =   5159
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   8730
            TabIndex        =   89
            Top             =   120
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   315
            Left            =   2055
            TabIndex        =   0
            Top             =   120
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton SearchCashCustomer 
            Height          =   375
            Index           =   0
            Left            =   8475
            TabIndex        =   109
            TabStop         =   0   'False
            ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
            Top             =   1200
            Width           =   375
            _ExtentX        =   661
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
            ButtonImage     =   "FrmReturnSalling33.frx":105C
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   660
            TabIndex        =   158
            Top             =   120
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   225
            Index           =   40
            Left            =   8010
            TabIndex        =   162
            Top             =   1200
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "VAT —ﬁ„"
            Height          =   225
            Index           =   39
            Left            =   8055
            TabIndex        =   161
            Top             =   870
            Width           =   465
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„·…"
            Height          =   285
            Index           =   65
            Left            =   930
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   150
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÌ"
            Height          =   345
            Index           =   37
            Left            =   10395
            TabIndex        =   131
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
            Height          =   285
            Index           =   35
            Left            =   13020
            TabIndex        =   111
            Top             =   1170
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·Ì›Ê‰"
            Height          =   345
            Index           =   84
            Left            =   13740
            TabIndex        =   110
            Top             =   1545
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„…"
            Height          =   300
            Index           =   34
            Left            =   1905
            TabIndex        =   107
            Top             =   570
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   315
            Index           =   24
            Left            =   3810
            TabIndex        =   106
            Top             =   525
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   55
            Left            =   720
            TabIndex        =   105
            Top             =   570
            Width           =   360
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»«∆⁄"
            Height          =   285
            Index           =   25
            Left            =   4305
            TabIndex        =   92
            Top             =   150
            Width           =   825
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   10260
            TabIndex        =   90
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Œ“Ì‰Â"
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   7275
            TabIndex        =   79
            Top             =   480
            Width           =   1320
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   7845
            TabIndex        =   75
            Top             =   120
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„·Ì…"
            Height          =   285
            Index           =   5
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   120
            Width           =   1710
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·⁄„·Ì…"
            Height          =   315
            Index           =   6
            Left            =   12570
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   465
            Width           =   1710
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   315
            Index           =   7
            Left            =   13020
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   285
            Index           =   8
            Left            =   13170
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   1920
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   315
            Index           =   9
            Left            =   9375
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   465
            Width           =   1725
         End
      End
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5160
         Left            =   15
         TabIndex        =   20
         Top             =   2985
         Width           =   14355
         _cx             =   25321
         _cy             =   9102
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
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…|«·ﬁÌ„… «·„÷«›…|FIFO"
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
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "FrmReturnSalling33.frx":1459
         Picture(1)      =   "FrmReturnSalling33.frx":17F3
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   4695
            Left            =   15600
            TabIndex        =   165
            TabStop         =   0   'False
            Top             =   45
            Width           =   14265
            _cx             =   25162
            _cy             =   8281
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
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   195
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   180
               Visible         =   0   'False
               Width           =   1200
            End
            Begin VB.CommandButton Command10 
               BackColor       =   &H8000000B&
               Caption         =   "«·€«¡ «·”œ«œ"
               Height          =   315
               Left            =   10920
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   120
               Width           =   1695
            End
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   3660
               Left            =   0
               TabIndex        =   169
               Top             =   600
               Width           =   14280
               _cx             =   25188
               _cy             =   6456
               Appearance      =   2
               BorderStyle     =   1
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
               Cols            =   18
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReturnSalling33.frx":1B8D
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
               ExplorerBar     =   1
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
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  ›Ê« Ì— «·„»Ì⁄«  ··⁄„Ì·"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   405
               Index           =   41
               Left            =   5880
               TabIndex        =   172
               Top             =   120
               Width           =   3015
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "≈Ã„«·Ì «·›Ê« Ì—"
               Height          =   255
               Left            =   12120
               RightToLeft     =   -1  'True
               TabIndex        =   171
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   4320
               Width           =   1575
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               BackColor       =   &H8000000B&
               BackStyle       =   0  'Transparent
               Height          =   375
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   170
               ToolTipText     =   "Ì „  Õ„Ì· Â–« «·„’—Ê› ⁄·Ï «·⁄„Ê·«  «·»‰ﬂÌ…"
               Top             =   4320
               Width           =   12135
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   4695
            Left            =   45
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   45
            Width           =   14265
            _cx             =   25162
            _cy             =   8281
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
            AutoSizeChildren=   8
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
            GridRows        =   5
            GridCols        =   5
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmReturnSalling33.frx":1E64
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   30
               TabIndex        =   68
               Top             =   4425
               Width           =   13155
               _ExtentX        =   23204
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
               Begin ImpulseButton.ISButton CmdDele 
                  Height          =   300
                  Left            =   3840
                  TabIndex        =   148
                  Top             =   0
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   529
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› «·„Õœœ"
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
                  ButtonImage     =   "FrmReturnSalling33.frx":1EEF
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   945
               Index           =   4
               Left            =   30
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   30
               Width           =   14205
               _cx             =   25056
               _cy             =   1667
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
               Begin VB.CheckBox chkTaxExempt 
                  Caption         =   "„⁄«›«…"
                  Height          =   315
                  Left            =   630
                  RightToLeft     =   -1  'True
                  TabIndex        =   176
                  Top             =   0
                  Width           =   945
               End
               Begin VB.TextBox TxtShortName 
                  Height          =   300
                  Left            =   4380
                  TabIndex        =   145
                  Top             =   0
                  Width           =   6870
               End
               Begin VB.TextBox TxtItemCodeB 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   11400
                  TabIndex        =   99
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   1635
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   5715
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Top             =   630
                  Width           =   1665
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   2250
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   630
                  Width           =   1380
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3630
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   630
                  Width           =   1965
               End
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   285
                  Left            =   780
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   630
                  Width           =   1320
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   7380
                  TabIndex        =   8
                  Top             =   630
                  Width           =   3825
                  _ExtentX        =   6747
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   11205
                  TabIndex        =   7
                  Top             =   630
                  Width           =   2880
                  _ExtentX        =   5080
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   345
                  Left            =   90
                  TabIndex        =   13
                  Top             =   585
                  Width           =   540
                  _ExtentX        =   953
                  _ExtentY        =   609
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
                  ButtonImage     =   "FrmReturnSalling33.frx":2489
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
               Begin ImpulseButton.ISButton SearchCashCustomer 
                  Height          =   360
                  Index           =   1
                  Left            =   11160
                  TabIndex        =   100
                  TabStop         =   0   'False
                  ToolTipText     =   "«÷€ÿ ·«÷«›… ⁄„Ì· ÃœÌœ"
                  Top             =   345
                  Visible         =   0   'False
                  Width           =   270
                  _ExtentX        =   476
                  _ExtentY        =   635
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
                  ButtonImage     =   "FrmReturnSalling33.frx":2823
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorShadow     =   -2147483631
                  ColorOutline    =   -2147483631
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·»ÕÀ «·”—Ì⁄"
                  Height          =   315
                  Index           =   97
                  Left            =   11625
                  TabIndex        =   146
                  Top             =   0
                  Width           =   1290
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   255
                  Index           =   31
                  Left            =   11775
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   345
                  Width           =   1425
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   255
                  Index           =   30
                  Left            =   7590
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   345
                  Width           =   3615
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   255
                  Index           =   29
                  Left            =   5805
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   345
                  Width           =   1575
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   255
                  Index           =   28
                  Left            =   3780
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   345
                  Width           =   1815
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   255
                  Index           =   27
                  Left            =   2460
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   345
                  Width           =   1170
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   255
                  Index           =   26
                  Left            =   780
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   345
                  Width           =   1320
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   3420
               Left            =   30
               TabIndex        =   144
               Top             =   990
               Width           =   14205
               _cx             =   25056
               _cy             =   6032
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
               Cols            =   20
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReturnSalling33.frx":2C20
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
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   240
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   4425
               Width           =   450
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4695
            Index           =   5
            Left            =   15000
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   45
            Width           =   14265
            _cx             =   25162
            _cy             =   8281
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
            AutoSizeChildren=   8
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
            GridRows        =   12
            GridCols        =   8
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmReturnSalling33.frx":2F70
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   4695
               Left            =   0
               TabIndex        =   101
               TabStop         =   0   'False
               Top             =   0
               Width           =   14265
               _cx             =   25162
               _cy             =   8281
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
               Begin VSFlex8UCtl.VSFlexGrid Grid 
                  Height          =   2730
                  Left            =   8760
                  TabIndex        =   102
                  Top             =   0
                  Width           =   5100
                  _cx             =   8996
                  _cy             =   4815
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
                  BackColor       =   -2147483640
                  ForeColor       =   65280
                  BackColorFixed  =   14871017
                  ForeColorFixed  =   -2147483630
                  BackColorSel    =   -2147483641
                  ForeColorSel    =   -2147483634
                  BackColorBkg    =   -2147483636
                  BackColorAlternate=   -2147483640
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
                  RowHeightMin    =   400
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmReturnSalling33.frx":3074
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
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Ìﬂ"
               Height          =   555
               Index           =   2
               Left            =   6135
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   2160
               Width           =   2595
            End
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "¬Ã· "
               Height          =   555
               Index           =   1
               Left            =   6135
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   2160
               Width           =   2595
            End
            Begin VB.CheckBox XPChkPayType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ﬁœ«"
               Height          =   360
               Index           =   0
               Left            =   6135
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   1155
               Width           =   2595
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   555
               Index           =   2
               Left            =   6135
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   2160
               Width           =   2595
               Begin VB.TextBox XPTxtChqueNum 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2970
                  MaxLength       =   40
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   75
                  Width           =   975
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Index           =   2
                  Left            =   2970
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   465
                  Width           =   975
               End
               Begin MSDataListLib.DataCombo DCboBankName 
                  Height          =   315
                  Left            =   60
                  TabIndex        =   38
                  Top             =   90
                  Width           =   1935
                  _ExtentX        =   3413
                  _ExtentY        =   556
                  _Version        =   393216
                  Locked          =   -1  'True
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSComCtl2.DTPicker XPDTPDueDate 
                  Height          =   345
                  Left            =   60
                  TabIndex        =   39
                  Top             =   465
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   609
                  _Version        =   393216
                  Format          =   147783681
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   19
                  Left            =   1620
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   525
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   210
                  Index           =   17
                  Left            =   2115
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   135
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   210
                  Index           =   18
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   142
                  Width           =   735
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   16
                  Left            =   4215
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   532
                  Width           =   465
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   555
               Index           =   1
               Left            =   6135
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   2160
               Width           =   2595
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   1320
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   90
                  Width           =   1305
               End
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   1
                  Left            =   3270
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   90
                  Width           =   1215
               End
               Begin MSComCtl2.DTPicker DtpDelayDate 
                  Height          =   330
                  Left            =   2160
                  TabIndex        =   31
                  Top             =   480
                  Width           =   1545
                  _ExtentX        =   2725
                  _ExtentY        =   582
                  _Version        =   393216
                  Format          =   147783681
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                  Height          =   210
                  Index           =   21
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   540
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   15
                  Left            =   4530
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   150
                  Width           =   405
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   14
                  Left            =   2670
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   150
                  Width           =   525
               End
            End
            Begin VB.Frame Fram 
               BackColor       =   &H00E2E9E9&
               BorderStyle     =   0  'None
               Height          =   480
               Index           =   0
               Left            =   4470
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   1035
               Width           =   1665
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   5820
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   120
                  Width           =   1215
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Index           =   0
                  Left            =   3870
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   120
                  Width           =   1305
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   270
                  Index           =   22
                  Left            =   3060
                  RightToLeft     =   -1  'True
                  TabIndex        =   61
                  Top             =   150
                  Width           =   765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   210
                  Index           =   13
                  Left            =   6990
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   180
                  Width           =   465
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   210
                  Index           =   12
                  Left            =   5220
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   150
                  Width           =   525
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   1035
               Index           =   20
               Left            =   6135
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   0
               Width           =   2595
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   4695
            Left            =   15300
            TabIndex        =   149
            TabStop         =   0   'False
            Top             =   45
            Width           =   14265
            _cx             =   25162
            _cy             =   8281
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
            Begin VB.TextBox txtManulaVat 
               Alignment       =   2  'Center
               BackColor       =   &H00C0FFFF&
               Height          =   375
               Left            =   120
               TabIndex        =   174
               Top             =   0
               Width           =   1215
            End
            Begin VB.CheckBox ChecVAT 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   225
               Left            =   12855
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   120
               Width           =   1065
            End
            Begin VB.TextBox TxtValueAdded 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   6945
               TabIndex        =   150
               Top             =   4335
               Width           =   2190
            End
            Begin VSFlex8UCtl.VSFlexGrid VatGrid 
               Height          =   3855
               Left            =   120
               TabIndex        =   151
               Tag             =   "1"
               Top             =   450
               Width           =   13965
               _cx             =   24633
               _cy             =   6800
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
               FormatString    =   $"FrmReturnSalling33.frx":3232
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«œŒ«· «·‰”»… «·ÌœÊÌ…"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   148
               Left            =   1320
               TabIndex        =   175
               Top             =   120
               Width           =   1800
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               Caption         =   "«’‰«› «·ﬁÌ„… «·„÷«›…"
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   10665
               TabIndex        =   153
               Top             =   570
               Width           =   3180
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   " «·«Ã„«·Ì"
               Height          =   240
               Index           =   104
               Left            =   9690
               TabIndex        =   152
               Top             =   4380
               Width           =   1110
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   615
         Left            =   15
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   15
         Width           =   14370
         _cx             =   25347
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
         Caption         =   "„—œÊœ«  «·„»Ì⁄« "
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
         Begin VB.TextBox TxtItemsIDes 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtPPointID 
            Height          =   285
            Left            =   0
            TabIndex        =   147
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3660
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   0
            Visible         =   0   'False
            Width           =   1395
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1905
            TabIndex        =   70
            Top             =   120
            Width           =   825
            _ExtentX        =   1455
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
            ButtonImage     =   "FrmReturnSalling33.frx":3346
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
            Left            =   1035
            TabIndex        =   71
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
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
            ButtonImage     =   "FrmReturnSalling33.frx":36E0
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
            Left            =   2805
            TabIndex        =   72
            Top             =   120
            Width           =   795
            _ExtentX        =   1402
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
            ButtonImage     =   "FrmReturnSalling33.frx":3A7A
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
            Left            =   120
            TabIndex        =   73
            Top             =   120
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "FrmReturnSalling33.frx":3E14
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
            Left            =   9810
            Picture         =   "FrmReturnSalling33.frx":41AE
            Stretch         =   -1  'True
            Top             =   120
            Width           =   525
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
            Index           =   11
            Left            =   3765
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   0
            Width           =   7035
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   465
         Index           =   3
         Left            =   0
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   8160
         Width           =   14355
         _cx             =   25321
         _cy             =   820
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
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   13695
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   -240
            Visible         =   0   'False
            Width           =   795
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   2985
            TabIndex        =   117
            Top             =   75
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblValueAdded 
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
            Left            =   7635
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   30
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·„÷«›…"
            Height          =   255
            Index           =   38
            Left            =   8640
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   30
            Width           =   1275
         End
         Begin VB.Label LBLGross 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """#,###.##"""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   5160
            TabIndex        =   143
            Top             =   840
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   210
            Left            =   8865
            RightToLeft     =   -1  'True
            TabIndex        =   130
            Top             =   105
            Visible         =   0   'False
            Width           =   75
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ì"
            Height          =   240
            Index           =   3
            Left            =   13050
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   75
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   0
            Left            =   2310
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   105
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   240
            Index           =   2
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   75
            Width           =   75
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   1380
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   75
            Width           =   705
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   75
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„” Œœ„"
            Height          =   285
            Index           =   1
            Left            =   4335
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   75
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   240
            Index           =   32
            Left            =   11115
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   75
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   240
            Index           =   33
            Left            =   6630
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   75
            Width           =   675
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
            Left            =   10110
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   0
            Width           =   930
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
            Height          =   330
            Left            =   11940
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   0
            Width           =   1035
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
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   0
            Width           =   1410
         End
         Begin VB.Label lblcost 
            Alignment       =   1  'Right Justify
            Caption         =   "Label2"
            Height          =   210
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   105
            Visible         =   0   'False
            Width           =   390
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   510
         Index           =   1
         Left            =   15
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   8640
         Width           =   14370
         _cx             =   25347
         _cy             =   900
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
         AutoSizeChildren=   8
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
         GridRows        =   1
         GridCols        =   20
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"FrmReturnSalling33.frx":7E16
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   0
            Left            =   12945
            TabIndex        =   134
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            ColorToggledText=   -2147483631
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   330
            Index           =   1
            Left            =   11498
            TabIndex        =   135
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   330
            Index           =   2
            Left            =   10057
            TabIndex        =   136
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   330
            Index           =   3
            Left            =   8616
            TabIndex        =   137
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   330
            Index           =   4
            Left            =   7175
            TabIndex        =   138
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   330
            Index           =   5
            Left            =   5734
            TabIndex        =   139
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   330
            Index           =   6
            Left            =   90
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   90
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
            Height          =   330
            Index           =   7
            Left            =   4290
            TabIndex        =   141
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   330
            Left            =   1411
            TabIndex        =   142
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
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
            Height          =   330
            Index           =   9
            Left            =   2852
            TabIndex        =   164
            Top             =   90
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… 2"
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   435
         Left            =   7185
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   7575
         Width           =   3555
      End
   End
End
Attribute VB_Name = "FrmReturnSalling33"
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
Dim FlgBillBuy As Boolean
Public BolPrint As Boolean
Dim bill_id As Double
Dim RsNotesGeneral  As ADODB.Recordset
Dim general_noteid  As Long
Dim SngTemp  As Single
Dim voucher_id As Double
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String
Dim DateChanged As Boolean
Dim TxtNoteSerial1V As String
Dim IsVouc         As Boolean
Private Sub Check1_Click()

    Dim i As Integer

    If Check1.value = vbChecked Then

        With Me.VSFlexGrid1
 
            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.VSFlexGrid1

            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If
    RelineBuy
End Sub
Sub RelineBuy()
    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.VSFlexGrid1
        For i = .FixedRows To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("RemainingValue")))
           End If
           Next i
  
    End With
   Label28.Caption = Sm
End Sub
Private Sub Command10_Click()
Dim i As Integer
Dim StrSQL As String
If Me.TxtModFlg.Text = "E" Then
DeleteBillBuy
VSFlexGrid1.Enabled = True
        Check1.Enabled = True
      StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtBillID.Text) & " and TransType=1"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtBillID.Text) & " and TransType=1"
    Cn.Execute StrSQL, , adExecuteNoRecords

            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
VSFlexGrid1.Rows = 1

FlgBillBuy = True
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «·€«¡ «·”œ«œ"
Else
MsgBox "Done"
End If
    With Me.VSFlexGrid1

            For i = .FixedRows To .Rows - 1
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i


        End With
End If
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 2106
        FrmCustemerSearch.show vbModal
     
    End If
    
End Sub

Private Sub Dcbranch_GotFocus()
Dcbranch_Click (0)
End Sub

Private Sub DcCurrency_Change()

    If Me.TxtModFlg.Text = "" Or Me.TxtModFlg.Text = "R" Then Exit Sub
    If Me.DcCurrency.BoundText <> "" Then
        txt_Currency_rate.Text = get_currency_rate(Me.DcCurrency.BoundText)
    Else
        txt_Currency_rate.Text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub
Public Sub RetriveSerials(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                          currentrow As Long, Optional Price As Double, Optional UnitID As Double = 1, Optional UnitName As String)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(strInputString, strFilterText)
    Dim i As Integer
 
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    
    Dim CurrentSerial As String
 
   
    '*****************************************************
    For intX = 0 To UBound(astrSplitItems)
   FG.Cell(flexcpData, Num, FG.ColIndex("Code")) = ItemID
   FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
   
        FG.TextMatrix(Num, FG.ColIndex("Name")) = ItemID
        
        
         FG.TextMatrix(Num, FG.ColIndex("UnitID")) = ItemID
        FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = 1
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = 0
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = 0
    
           'FG.TextMatrix(I, FG.ColIndex("HaveSerial")) = True
         
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
        
        FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = UnitID
        FG.TextMatrix(Num, FG.ColIndex("UnitID")) = UnitName
             FG.TextMatrix(i, FG.ColIndex("HaveSerial")) = True
             
        
If val(Price) > 0 Then
            FG.TextMatrix(Num, FG.ColIndex("price")) = Price
        End If
        
        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.Rows = FG.Rows + 1
 
        Num = Num + 1
    If intX = UBound(astrSplitItems) Then
    NewGrid.Calculate Num
    NewGrid.bassprofit = True
    NewGrid.DtpBillDate_Change
        End If
    Next
     
     
    TxtFillData.Text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub


Public Sub RetriveSerialsx(ItemID As String, _
                          ItemName As String, _
                          seriallist As String, _
                          currentrow As Long)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(strInputString, strFilterText)
    Dim i As Integer
    ' For i = 1 To Fg.Rows - 2
    '        If Fg.TextMatrix(i, Fg.ColIndex("Code")) = ItemID Then
    '         Me.Fg.RemoveItem (i)
    '         i = 1
    '        End If
    'NewGrid.Grid_AfterEdit Num, Fg.ColIndex("Code")
    ' Next i
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    For intX = 0 To UBound(astrSplitItems)
   
        FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
        NewGrid.Grid_AfterEdit Num, FG.ColIndex("Code")
        ' FG.TextMatrix(Num, FG.ColIndex("Name")) = itemname
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
  
        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.Rows = FG.Rows + 1
 
        Num = Num + 1
    Next
 
    TxtFillData.Text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub
 Function CuurentLogdata(Optional Currentmode As String)
   
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·›« Ê—…   " & TxtNoteSerial1.Text & CHR(13) & " «· «—ÌŒ " & XPDtbBill.value & CHR(13) & " «·Œ“Ì‰… " & DcboBox.Text & CHR(13) & " «·„Œ“‰  " & DCboStoreName.Text & CHR(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.Text & CHR(13) & "‰Ê⁄ «·”‰œ " & DCDocTypes & CHR(13) & "ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & CHR(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Bill No " & TxtNoteSerial1.Text & CHR(13) & " Date " & XPDtbBill.value & CHR(13) & " Box " & DcboBox.Text & CHR(13) & " Store  " & DCboStoreName.Text & CHR(13) & " Supplier/Cuxtomer" & DBCboClientName.Text & CHR(13) & "Doc Type" & DCDocTypes & CHR(13) & "Payment Type" & CboPayMentType & CHR(13) & " GE NO" & TxtNoteSerial
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 170, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 170, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function
Function SaveItemsData(Optional Transaction_ID As String = 0)

If SystemOptions.WorkWithItemsDetails = False Then Exit Function
       Dim RsgGrantee    As New ADODB.Recordset
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim AllDes As String
    Dim RowNum As Integer
    Dim StrSQL As String
    strFilterText = ","
    Set RsgGrantee = New ADODB.Recordset
    Cn.Execute "delete ItemsDetails   where Transaction_ID= " & (Me.XPTxtBillID.Text)
    
  '  RsgGrantee.Open "TBLRegularMaint", Cn, adOpenStatic, adLockOptimistic, adCmdTable

   StrSQL = "SELECT    * from  ItemsDetails Where (1 = -1)"
   RsgGrantee.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
     
 
    Dim strFilterText1 As String
      Dim UnitName As String
    Dim ttypename As String
     Dim typename As String
 
 
 
 
    Dim inty As Integer
    Dim intervalstr As String
Dim Name As String
Dim NameE As String
Dim Remarks As String
Dim NooFRows As Double
    
     Dim astrSplitItems1() As String
 
    strFilterText = "&&"
         strFilterText1 = "@@"
     
    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            
           If FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) <> "" Then
                AllDes = FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea"))
                astrSplitItems = Split(AllDes, strFilterText)
         NooFRows = UBound(astrSplitItems) + 1
                For intX = 0 To NooFRows - 2
             
                
                          RsgGrantee.AddNew
                         astrSplitItems1 = Split(astrSplitItems(intX), strFilterText1)
                         RsgGrantee("ItemDetailedCode").value = (astrSplitItems1(0))
                         RsgGrantee("ParrtNoCode").value = (astrSplitItems1(1))
                         RsgGrantee("count").value = val(astrSplitItems1(2))
                         RsgGrantee("unitid").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 1, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))  ' val(astrSplitItems1(3))
                         RsgGrantee("ColorID").value = val(astrSplitItems1(4))
                         RsgGrantee("sizeid").value = val(astrSplitItems1(5))
                         RsgGrantee("ClassId").value = val(astrSplitItems1(6))
                         RsgGrantee("ProductionDate").value = IIf(IsDate((astrSplitItems1(7))), astrSplitItems1(7), Null)
                         RsgGrantee("ExpireDate").value = IIf(IsDate((astrSplitItems1(8))), astrSplitItems1(8), Null)
                        RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
                        RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
                       RsgGrantee("EffectN").value = 1
                    RsgGrantee.update
                                    Next intX
                Else
                If FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) <> "" Then
                RsgGrantee.AddNew
              RsgGrantee("ParrtNoCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))
            RsgGrantee("count").value = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
            RsgGrantee("unitid").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
          RsgGrantee("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RsgGrantee("sizeid").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RsgGrantee("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            RsgGrantee("Transaction_ID").value = val(Me.XPTxtBillID.Text)
           RsgGrantee("ItemId").value = FG.TextMatrix(RowNum, FG.ColIndex("Code"))
          RsgGrantee("ItemDetailedCode").value = FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))
          RsgGrantee("EffectN").value = 1
           RsgGrantee.update
                  
         End If
         
                  
                   
                   End If
                   

 
                
  
                    
            End If

       

    Next RowNum


End Function

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer)
Dim usedaccount As Integer
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim LngCurItemID As Double
    Dim LngUnitID As Long
    Dim UnitFactor As Double

    Dim TOTAL_COST As Variant

    With FG

        For i = 1 To FG.Rows - 1
            LngCurItemID = val(FG.TextMatrix(i, FG.ColIndex("Code")))
            LngUnitID = val(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")))
            
            GetUnitNoOfItems LngCurItemID, LngUnitID, UnitFactor

            If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("itemtype"))) <> 1 Then
                TOTAL_COST = TOTAL_COST + (FG.TextMatrix(i, FG.ColIndex("Count")) * FG.TextMatrix(i, FG.ColIndex("ItemCostPrice")))
            End If

        Next i

    End With

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) + val(TxtValueAdded.Text)
    my_branch = BranchID

    If TOTAL_COST > 0 Then
   
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            StrTempDes = "”‰œ «÷«›Â  —ﬁ„ " & Me.TxtNoteSerial1.Text & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄« "
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.Text
            Else
                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.Text
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.Text
                        Else
                            StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.Text
                        End If
            
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) + val(TxtValueAdded.Text)

        If TOTAL_COST > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(1, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ  ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

         '       StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄« 
            
            
              If val(DCDocTypes.BoundText) > 0 Then '„—œÊœ«  «·„»Ì⁄« 
                getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

                        If StrTempAccountCode = "" And usedaccount = 1 Then
                                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ  «·«” ·«„ ", vbCritical
                                    GoTo ErrTrap
                        ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                        
                        ElseIf usedaccount = 0 Then
                                StrTempAccountCode = Account_Code_dynamic '
                        End If

            Else
                        StrTempAccountCode = Account_Code_dynamic '
          End If
          
          
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.Text
                Else
                    StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TOTAL_COST, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.Rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 1)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»    ﬂ·›… «·„»Ì⁄«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(i, FG.ColIndex("Code")), 1, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ    «÷«›Â —ﬁ„ " & TxtNoteSerial1V & "»‰«¡ ⁄·Ï „—œÊœ«  „»Ì⁄«  »—ﬁ„ " & Me.TxtNoteSerial1.Text
                            Else
                                StrTempDes = "Issue Voucher No. " & TxtNoteSerial1V
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If
    End If

    Dim StrSQL  As String
    StrSQL = "UPDATE Transactions SET NOTS=" & val(Me.Text1.Text) & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.Text)
    'sql = "update transactions set closed=1" & ",nots=" & Val(Me.XPTxtBillID.text) & ",nots2=" & Me.TxtNoteSerial1.text & " where  Transaction_ID= " & Val(Me.Text1.text)
    Cn.Execute StrSQL
    updateNotesValueAndNobytext CDbl(general_noteid)
    
ErrTrap:
End Function

Private Function CreateRecieveVoucher() As Boolean
     On Error GoTo ErrTrap
IsVouc = False
CreateRecieveVoucher = False
    Dim UnitID As Long
    Dim i As Long

    If CboRetrunType.ListIndex = 1 Then

        With FG

            For i = 1 To FG.Rows - 1

                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                    UnitID = IIf(FG.Cell(flexcpData, i, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, i, FG.ColIndex("UnitID"))))

    '                If val(fg.TextMatrix(i, fg.ColIndex("ItemCostPrice"))) = 0 Then
                                               
     
    '                    If SystemOptions.UserInterface = ArabicInterface Then
    '                        MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ  ﬂ·›Â «·»Ì⁄ ·Â Ê·„ Ì „  ÕœÌœ À„‰ «·‘—«¡ Ê·Ì” ·Â ﬁÌ„Â —’Ìœ «›  «ÕÌ… ·–·ﬂ ·« Ì„ﬂ‰ «‰‘«¡ ”‰œ «·«÷«›Â "
    '                    Else
    '                        MsgBox "Item in line no " & i & "Group Name Account Not Defined"
    '                    End If

    '                    Exit Sub
    '                End If
                End If

            Next i

        End With

    End If

    Dim groupAccount  As String

    If detect_inventory_work_type = 3 Then
   
        With FG

            For i = 1 To FG.Rows - 1

                If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
                
                    ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                    groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                    If groupAccount = "Error" Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                        Else
                            MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                        End If

                        Exit Function
                    End If
                End If

            Next i

        End With

    End If

    'CurrentVoucherNo = GetVoucherGLNO(Val(Text1.text))
    '  DeleteTransactiomsVoucher val(Text1.text)
    '   Dim RowNum As Integer

    '   For RowNum = 1 To Fg.Rows - 1
    '                    If Fg.TextMatrix(RowNum, Fg.ColIndex("Code")) <> "" Then
    '
    '                     If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
    '
    ''
    '                 Else '€Ì— „ﬁÌœ »›« Ê—…
    '                     unitid = IIf(Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID")) = "", Null, (Fg.Cell(flexcpData, RowNum, Fg.ColIndex("UnitID"))))
    '                    Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCostPrice")) = ModItemCostPrice.GetCostItemPrice(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, , unitid)
    '
    '                 End If
    '
    '            End If
 
    'Next RowNum
 
    Dim MYWAER As String
    Dim StrSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
 
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
 
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim MYTEXT As String
    '>>>>>>>>>>>>>>>>>>>>>>>>>

    rs.Close
 
    Dim xyeas As Boolean
    xyeas = True

    If xyeas = True Then
 
        MYTEXT = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
        'mytext = TxtTransSerial.text

        '         rs!nots = mytext
        '         rs.update

        Dim Transaction_ID As Long
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        Dim general_noteid As Long
        Dim RsNotesGeneral As ADODB.Recordset
        Dim TxtNoteSerialV As String
            
        my_branch = Me.dcBranch.BoundText

        If TxtNoteSerialV = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Function
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Function
                Else
                    TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If

        If TxtNoteSerial1V = "" Then
        TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , val(DCboStoreName.BoundText), , , , val(DCboUserName.BoundText))
        
            If TxtNoteSerial1V = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  «÷«›Â ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Function
            Else
                       
                If TxtNoteSerial1V = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·«” ·«„ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
                Else
                 '   TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20)
                End If
            End If
        End If
                 
        If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
            TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
            TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
        End If
           
        Dim sql As String
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
        Text1.Text = Transaction_ID

        sql = "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,nots2,NoteSerial,NoteSerial1,NoteId,BranchId,Closed,ManualNO,CBoBasedON)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 20,CusID,StoreID,UserID,Emp_ID,nots=" & val(XPTxtBillID.Text) & ",nots2='" & TxtNoteSerial1.Text & "' ,NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId,1,ManualNO ,12 From Transactions Where Transaction_ID =" & val(XPTxtBillID.Text) & " And Transaction_Type =9"
      
        Cn.Execute sql

        
    sql = "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,ClassId,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ProductionDate,ExpiryDate,LotNO,OldQty,OldCost,NewQty,NewCost)"
 sql = sql & "      SELECT costprice,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, price , ColorID,ItemSize,ClassId,UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ProductionDate,ExpiryDate,LotNO ,OldQty,OldCost,NewQty,NewCost From dbo.Transaction_Details Where   Transaction_ID = " & XPTxtBillID.Text
    sql = sql & "  AND  Item_ID NOT IN ("
  sql = sql & "  SELECT     dbo.Transaction_Details.Item_ID  FROM         dbo.Transaction_Details INNER JOIN                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
sql = sql & "  Where (dbo.TblItems.ItemType = 1)"
 sql = sql & "  And  dbo.Transaction_Details.Transaction_ID = " & val(XPTxtBillID.Text) & ")"
  

 

          Cn.Execute sql
             UpdateTransactionsCost CStr(Transaction_ID)
             
        Text1.Text = Transaction_ID
        'TxtIssueSerial.text = TxtNoteSerial1V
        'Create big notes
     
        Set RsNotesGeneral = New ADODB.Recordset
'        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

 StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


        If Me.TxtModFlg.Text = "N" Then
        Else
            general_noteid = val(TxtNoteID.Text)
        End If
        
        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TxtNoteID.Text = general_noteid
        RsNotesGeneral("Transaction_ID").value = Transaction_ID
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 160
        RsNotesGeneral("Note_Value").value = Null
        RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
        'RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        RsNotesGeneral("Remark").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
        
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(9) '«–‰ wvt
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

        CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)
       
        sql = " update Transactions Set NoteID = " & general_noteid & " where Transaction_ID = " & Transaction_ID
        Cn.Execute sql
    End If
    IsVouc = True
    CreateRecieveVoucher = True
 Exit Function
    '
 
ErrTrap:
    IsVouc = True
    CreateRecieveVoucher = False
    Exit Function

End Function

Private Sub C1Elastic6_DblClick()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboRetrunType_Change()

    If CboRetrunType.ListIndex = 0 Then
        lbl(4).Enabled = True
        'Me.TxtInvSerial.Enabled = True
        Me.CmdOpenTrans.Enabled = True
        Me.CmdSearchTrans.Enabled = True
    ElseIf Me.CboRetrunType.ListIndex = 1 Then
        lbl(4).Enabled = False
        'Me.TxtInvSerial.Enabled = False
        Me.CmdOpenTrans.Enabled = False
        Me.CmdSearchTrans.Enabled = False
    End If
    If Me.TxtModFlg.Text <> "R" Then
                If val(CboRetrunType.ListIndex) = 0 Then
                            NewGrid.ReturnTyp = 1 '2
                             VatGrid.Clear flexClearScrollable, flexClearEverything
                                VatGrid.Rows = 1
                 Else
                           NewGrid.ReturnTyp = 1
                              VatGrid.Clear flexClearScrollable, flexClearEverything
                                VatGrid.Rows = 1
                 End If
   NewGrid.DtpBillDate_Change
NewGrid.Calculate 1, , , True
 End If
 If SystemOptions.AllowReturnFIFO = True Then
BillCustomer
End If
End Sub
Private Sub ChecVAT_Click()
  Dim i As Integer
If Me.TxtModFlg.Text <> "R" Then
    If ChecVAT.value = vbChecked Then

        With Me.VatGrid
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = True
            Next i

        End With

    Else

        With Me.VatGrid

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("select")) = False
            Next i

        End With

    End If
    RelinVatGrid
    End If
End Sub
Private Sub CboRetrunType_Click()
    CboRetrunType_Change
End Sub

Function checkretutn() As Boolean
    Dim Msg As String
    checkretutn = False

    If Not IsDate(txtInvDate) Then Exit Function
    '    If SystemOptions.ReturnSallingOption = True Then
    Dim NoofDays As Integer

    If Me.TxtModFlg = "R" Or Me.TxtModFlg = "" Then Exit Function
    NoofDays = DateDiff("d", IIf(IsDate(Me.txtInvDate.Text), Me.txtInvDate.Text, Date), Me.XPDtbBill.value)
 
    If opt(0).value = True Then
        If NoofDays > SystemOptions.ReturnSallingIntervalCount Then
            Msg = " ·« Ì„ﬂ‰ «—Ã«⁄ Â–… «·›« Ê—… ·«‰ «·Õœ «·«ﬁ’Ï ··«—Ã«⁄ " & SystemOptions.ReturnSallingIntervalCount & "  ÌÊ„ " & CHR(13)
            Msg = Msg & " «·›« Ê—Â „‰  " & NoofDays & "  ÌÊ„ "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            checkretutn = False
            Exit Function
        
        End If
  
    Else

        If NoofDays > SystemOptions.ReturnSallingIntervalCount1 Then
            Msg = " ·« Ì„ﬂ‰ «” »œ«·  Â–… «·›« Ê—… ·«‰ «·Õœ «·«ﬁ’Ï ··«” »œ«· " & SystemOptions.ReturnSallingIntervalCount1 & "  ÌÊ„ " & CHR(13)
            Msg = Msg & " «·›« Ê—Â „‰  " & NoofDays & "  ÌÊ„ "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            checkretutn = False
            Exit Function
        End If

    End If
   
    checkretutn = True
         
    'End If
End Function

Public Sub Cmd_Click(Index As Integer)
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTest As ADODB.Recordset
  ' On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            clear_all Me
            TxtModFlg.Text = "N"
            VatGrid.Clear flexClearScrollable, flexClearEverything
            VatGrid.Rows = 1
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=9"))
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = val(GetSetting(StrAppRegPath, "DefaultOptions", "DefaultClient", 2))
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSaleStore", 1)
          '  DCboStoreName.BoundText = intDef
            Me.DcboBox.BoundText = 1
            XPTab301.CurrTab = 0
            DcCurrency.BoundText = 1
            FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                             If SystemOptions.BranchCanNotEdit = True Then
                             Me.dcBranch.Enabled = False
                             Else
                               Me.dcBranch.Enabled = True
                             End If
                                    Else
                               If SystemOptions.BranchCanNotEdit = True Then
                             Me.dcBranch.Enabled = False
                             Else
                               Me.dcBranch.Enabled = True
                             End If
                              End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = True
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If



            Me.dcBranch.BoundText = Current_branch
                  
            Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim EmpID As Integer
           Dim userbranchid As Integer
           Dim CUSTID As Integer
            'GetBranchData branch_id, dstore, dBox
            If Voucher_coding(val(my_branch), XPDtbBill.value, 14, 220, , , , , , , , val(DCboUserName.BoundText)) = "" And val(my_branch) <> 0 Then
                TxtNoteSerial1.locked = False
            Else
                TxtNoteSerial1.locked = True
 
            End If

            GetUserData user_id, usertype, userbranchid, dstore, dBox, , EmpID, , CUSTID
           
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
                  If SystemOptions.BranchCanNotEdit = True Then
                             Me.dcBranch.Enabled = False
                             Else
                               Me.dcBranch.Enabled = True
                             End If
                            
             '   DcboBox.Enabled = False
                DCboStoreName.Enabled = True
                Me.dcBranch.BoundText = userbranchid
                Me.DCboStoreName.BoundText = dstore
                Me.DcboBox.BoundText = dBox
                Me.DBCboClientName.BoundText = CUSTID
                 Me.DcboEmp.BoundText = EmpID
                 
            Else
                   If SystemOptions.BranchCanNotEdit = True Then
                             Me.dcBranch.Enabled = False
                             Else
                               Me.dcBranch.Enabled = True
                             End If
                DcboBox.Enabled = True
                DCboStoreName.Enabled = True
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
                Me.DcboBox.BoundText = ""
            End If

            If SystemOptions.ReturnSallingOption = True Then

                CboRetrunType.ListIndex = 0
                CboRetrunType.Enabled = False
            End If

            opt(1).value = True

            If Current_branch = 0 Then
                branch_id = Current_branch
                Me.dcBranch.BoundText = Current_branch
            End If
 FillGridWithData
 DcboEmp.SetFocus
CboRetrunType.ListIndex = 0

        Case 1
            If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

    '        If SystemOptions.usertype = UserNormal Then
    '            Msg = "·Ì” ·ﬂ Õﬁ  ⁄œÌ· ›Ï «·›Ê« Ì—"
    '            MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
    '            Exit Sub
    '        End If
    TxtModFlg.Text = "E"
 If SystemOptions.AllowReturnFIFO = True And val(DBCboClientName.BoundText) <> 0 And val(CboRetrunType.ListIndex) = 1 Then
Command10_Click
BillCustomer
End If
            
            Me.DCboUserName.BoundText = user_id

            If SystemOptions.ReturnSallingOption = True Then

                CboRetrunType.ListIndex = 0
                CboRetrunType.Enabled = False
            End If
If val(CboRetrunType.ListIndex) = 0 Then
 NewGrid.ReturnTyp = 2
 Else
 NewGrid.ReturnTyp = 1
 End If
            CuurentLogdata
DcboEmp.SetFocus
        Case 2
                             If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·« "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
 
 If CboRetrunType.ListIndex = 0 Then
 If val(TxtInvID.Text) = 0 Then
 MsgBox "Õœœ «·›« Ê—…", vbCritical
 Exit Sub
End If
 
 End If
            If SystemOptions.USERautoIssueVoucher = False And CboRetrunType.ListIndex = 0 And SystemOptions.returnnotcreatvoucher = False And SystemOptions.AllowReturnWithoutCost = False Then

                bill_id = val(TxtInvID.Text)
                
                voucher_id = check_bill_voucher(bill_id, 19) '·«ÌÃ«œ —ﬁ„ «–‰ «·’—› „‰ ﬁ«⁄œ… «·»Ì«‰« 

                If voucher_id = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·« ÌÊÃœ ”‰œ ’—› „Œ“‰Ì ·Â–… «·›« Ê—… Õ Ï Ì„ﬂ‰ Õ”«»  ﬂ·›… «·„»Ì⁄« ", vbCritical
                    Else
                        MsgBox " There is no issue voucher to this bill ", vbCritical
                    End If

                    GoTo ErrTrap
                End If
                   
                If checkretutn = False Then

                    Exit Sub
                End If

            End If

            Set RsNotesGeneral = New ADODB.Recordset
         '   RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
          
           StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            my_branch = Me.dcBranch.BoundText
      
            '            If Me.TxtModFlg.text = "N" Then
             
            '             End If
  If CheckFilegrid() = True Then
  If val(Me.TxtValueAdded.Text) > 0 Then
If GetValueAddedAccount(XPDtbBill.value, , , 1, 9) = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·ﬁÌ„… «·„÷«›…"
Else
MsgBox "Value added account not specified"
End If
Exit Sub
End If
End If
If SystemOptions.AllowReturnFIFO = True Then
BillCustomer
AutoCalculate
End If

            SaveData
     End If
        Case 3
            Call Undo

        Case 4
                             If ChekClodePeriod(XPDtbBill.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
                  
                  
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

    '        If SystemOptions.usertype = UserNormal Then
    '            Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
    '            MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
    '            Exit Sub
    '        End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

            FrmBuySearch.DealingForm = ReturnSalling
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ „—œÊœ«  «·„»Ì⁄« "
            FrmBuySearch.show vbModal

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport 0
        
        Case 9

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            PrintReport 1

        Case 6
            Unload Me

        Case 10
            ShowGL_cc TxtNoteSerial.Text, , 200, val(Me.TxtNoteID.Text)

        Case 8
       
        FrmInpout.XPBtnMove_Click (2)
            FrmInpout.Retrive val(Me.Text1.Text)
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdDele_Click()
 Dim i As Integer
With FG
i = .Rows - 1
Do While i > 0
If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
.RemoveItem i
End If
i = i - 1
Loop
End With
NewGrid.DtpBillDate_Change
NewGrid.Calculate 1, , , True
End Sub

Private Sub CmdHelp_Click()
'    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
'    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments Me.TxtNoteSerial1, "0903201701"

End Sub

Public Sub CmdOpenTrans_Click()
    Dim Msg As String
    Dim FrmNewSales As frmsalebill

    If val(Me.TxtInvSerial.Text) = 0 Then
      '  Msg = "»—Ã«¡ ﬂ «»… —ﬁ„ «·›« Ê—… ·Ì „ ⁄—÷Â«..!!"
      '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Dim Transaction_ID As Long
    Dim Transaction_Date  As Date
    'Me.TxtInvID.text = GetTransIDSerial(0, , Trim(Me.TxtInvSerial.text), "2 Or Transaction_Type = 21", , TxtInvID.text)
    GetTransIDFromNoteSerial1 Me.TxtInvSerial.Text, Transaction_ID, Transaction_Date, 21
    Me.TxtInvID.Text = Transaction_ID
    Me.txtInvDate.Text = Transaction_Date

    If val(Me.TxtInvID.Text) = 0 Then
        Msg = "·« ÊÃœ ›« Ê—… »Â–« «·—ﬁ„ ..!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        Retrive_Sales_invoice_data (val(Me.TxtInvID.Text))
        RetriveQtyItem (val(Me.TxtInvID.Text))
        RetriveValueAddedData val(Me.TxtInvID.Text)


    End If
   NewGrid.DtpBillDate_Change
NewGrid.Calculate 1, , , True
End Sub

Private Sub CmdSearchTrans_Click()
    ' ›« Ê—… „»Ì⁄« 
    'Load FrmBuySearch
    'FrmBuySearch.DealingForm = InvoiceTransaction
    'Set FrmBuySearch.ExtraRetrunObject = Me.TxtInvID
    'Set FrmBuySearch.ExtraRetrunObject1 = Me.TxtInvSerial
    'Set FrmBuySearch.ExtraRetrunObject2 = Me.txtInvDate
    '
    'FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ⁄„·Ì… »Ì⁄"
    'FrmBuySearch.DCboClientsName.BoundText = Me.DBCboClientName.BoundText
    'FrmBuySearch.Show vbModal
    
    
    
       FrmBuySearch.DealingForm = GridTransType.InvoiceTransaction
     FrmBuySearch.Index = 13
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „»Ì⁄«    "
            FrmBuySearch.show vbModal
            
 


End Sub

 Function Retrive_Items_data1()
    Dim StrSQL  As String
    Dim row_count As Long
    Dim Num As Long
    Dim i As Long
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    StrSQL = "select * from TblItems where ItemID in(" & TxtItemsIDes.Text & ")"
    rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If rs2.RecordCount > 0 Then
        
        If FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Code")) = "" Then
      FG.Rows = FG.Rows - 1
        End If
     With FG
     row_count = FG.Rows
       rs2.MoveFirst
       .Rows = rs2.RecordCount + .Rows
        For Num = row_count To .Rows - 1 'RsDetails.RecordCount
        .TextMatrix(Num, .ColIndex("Code")) = IIf(IsNull(rs2("ItemID").value), 0, rs2("ItemID").value)
      
        rs2.MoveNext
        Next Num
        For i = row_count To .Rows - 1 'RsDetails.RecordCount
          NewGrid.Grid_AfterEdit i, .ColIndex("Code")
        Next i
        NewGrid.Grid_AfterEdit row_count, .ColIndex("Code")
    End With
    End If


End Function

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then

        Load FrmItemSearch
        FrmItemSearch.RetrunType = 9
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub DCboStoreName_Change()
 TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

     If CheckStoreCoding(val(dcBranch.BoundText), 14) = True Or CheckStoreCoding(val(dcBranch.BoundText), 9) = True Then
     TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
        TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    CurrentVoucherNo = ""
TxtNoteSerial1V = ""
    DateChanged = True
    

     End If
     
    End If

End Sub

Private Sub Dcbranch_Change()
Dim Dcombos As New ClsDataCombos
 
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then

       Dcombos.GetDocTypebyid Me.DCDocTypes, 9, val(Me.dcBranch.BoundText)
       
    DateChanged = True

    End If
End Sub

Private Sub Dcbranch_Click(Area As Integer)
    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    
        TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    CurrentVoucherNo = ""
TxtNoteSerial1V = ""
    DateChanged = True
    
    
End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial), 240
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , val(Me.TxtNoteSerial), val(Me.TxtNoteSerial1), 240
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.Name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), val(Me.TxtNoteSerial), Me.TxtNoteSerial1, 240

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub Fg_DblClick()
    'FrmItemsDetails.Show
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption



End Sub



Private Sub SearchCashCustomer_Click(Index As Integer)
Select Case Index
Case 0
frmCashCustomerSearch.RetrunType = 3
frmCashCustomerSearch.show
Case 1
        Load FrmItemSearch2
        FrmItemSearch2.RetrunType = 2
        FrmItemSearch2.show vbModal
End Select

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.Text = "F" Then
        NewGrid.Calculate 1
    End If
RelinVatGrid
End Sub

Private Sub TxtInvID_Change()
    Dim Msg  As String

    If Me.TxtModFlg = "R" Then Exit Sub
    If val(Me.TxtInvID.Text) = 0 And Me.CboRetrunType.ListIndex = 0 Then
        '    Msg = "·« ÊÃœ ›« Ê—… »Â–« «·—ﬁ„ ..!!"
        '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

        'FrmReturnSalling.Retrive Val(Me.TxtInvID.text)
        ' Retrive_Sales_invoice_data (Val(Me.TxtInvID.text))
        If checkretutn = False Then
            Exit Sub
        End If
 
    End If

End Sub


Private Sub TxtInvSerial_KeyDown(KeyCode As Integer, _
                                 Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtInvDate.Text = ""
        CmdOpenTrans_Click
       NewGrid.Calculate 1, , , True
        NewGrid.DtpBillDate_Change
        
   
 
 
    End If
BillCustomer
   ' If KeyCode = vbKeyReturn Then
   '     txtInvDate.text = ""
   '     CmdOpenTrans_Click
   '      NewGrid.Calculate 1, , , True
   ' End If

End Sub

Private Sub TxtInvSerial_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then

           FrmBuySearch.DealingForm = GridTransType.InvoiceTransaction
     FrmBuySearch.Index = 13
            FrmBuySearch.Caption = "«·»ÕÀ ⁄‰ ›« Ê—… „»Ì⁄«    "
            FrmBuySearch.show vbModal
            
End If

   BillCustomer
End Sub

Private Sub TxtItemCodeB_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Load FrmItemSearch2
        FrmItemSearch2.RetrunType = 2
        FrmItemSearch2.show vbModal
  End If
  
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub



Private Sub TxtShortName_KeyDown(KeyCode As Integer, Shift As Integer)
SerchItems (TxtShortName.Text)
DoEvents
DoEvents
DoEvents
DoEvents

        If KeyCode = vbKeyReturn Then
        
        
   DCboItemsName.SetFocus
   DCboItemsName.BoundText = ""
         SendKeys "{F4}"
        End If
End Sub
Sub SerchItems(Optional str As String)
 
Dim sql As String
Dim SQL1 As String
   
    SerchItemspUBLIC str, sql, SQL1
    fill_combo DCboItemsCode, sql
  fill_combo DCboItemsName, SQL1
        
         
End Sub

Sub SerchItemsxx(Optional str As String)
 
Set DCboItemsCode.RowSource = Nothing
Set DCboItemsName.RowSource = Nothing
If str <> "" Then
Dim sql As String
Dim SQL1 As String
 
Dim StrWhere As String
  Dim astrSplit2tems2() As String
  Dim j As Integer
  Dim nElements As Integer
  Dim SearchString As String
StrWhere = ""
SearchString = ""
sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
If SystemOptions.UserInterface = ArabicInterface Then
SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If

          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
          If nElements = 0 Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                            StrWhere = " and (ItemName Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%'  or fullcode Like N'%" & Trim(str) & "%') "
                    Else
                            StrWhere = " and (ItemNamee Like N'%" & Trim(str) & "%' or barCodeNO Like N'%" & Trim(str) & "%' or shortName Like N'%" & Trim(str) & "%' or fullcode Like N'%" & Trim(str) & "%' ) "
                    End If
                    
          End If
        If nElements > 0 Then
        
     '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        SearchString = ""
        For j = 0 To nElements
        
         SearchString = SearchString & "%" & Trim(astrSplit2tems2(j))
             '     SearchString = "%" & Trim(astrSplit2tems2(j)) & SearchString
                  
        '   StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
        '   StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         SearchString = SearchString & "%"
                             If SystemOptions.UserInterface = ArabicInterface Then

             StrWhere = StrWhere + " and (ItemName Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             Else
              StrWhere = StrWhere + " and (ItemNamee Like '" & SearchString & "' or barCodeNO Like '" & SearchString & "' or shortName Like '" & SearchString & "') "
             End If
        '-  StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
      
         End If
        
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemName "
    End If


    SQL1 = SQL1 & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        SQL1 = SQL1 + " Order BY ItemName "
    Else
        SQL1 = SQL1 + " Order BY ItemNamee "
    End If
    
   End If
    fill_combo DCboItemsCode, sql
        fill_combo DCboItemsName, SQL1
        DoEvents
        DoEvents
                  If str = "" Then
                                 sql = " select  ItemID,barCodeNO   from  dbo.TblItems where TblItems.IsArchive=0"
                                 If SystemOptions.UserInterface = ArabicInterface Then
                                 SQL1 = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
                                     SQL1 = SQL1 + " Order BY ItemName "
                                 Else
                                 SQL1 = " select  ItemID,ItemNamee   from  dbo.TblItems where  TblItems.IsArchive=0 "
                                     SQL1 = SQL1 + " Order BY ItemNameE "
                                 End If
                                 
                                     fill_combo DCboItemsCode, sql
                                         fill_combo DCboItemsName, SQL1
                End If
        
       Exit Sub
       
If str <> "" Then
'Dim Sql As String
'Dim StrWhere As String
'  Dim astrSplit2tems2() As String
'  Dim j As Integer
'  Dim nElements As Integer
StrWhere = ""
If SystemOptions.UserInterface = ArabicInterface Then
sql = " select  ItemID,ItemName   from  dbo.TblItems where TblItems.IsArchive=0"
Else
sql = " select  ItemID,ItemNamee   from  dbo.TblItems where TblItems.IsArchive=0"
End If
          astrSplit2tems2 = Split(str, " ")
          nElements = UBound(astrSplit2tems2) - LBound(astrSplit2tems2)
        If nElements > 0 Then
        StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(0)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(0)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(0)) & "%') "
        For j = 1 To nElements - 1
        
           StrWhere = StrWhere + " and (ItemName Like N'%" & Trim(astrSplit2tems2(j)) & "%' or barCodeNO Like N'%" & Trim(astrSplit2tems2(j)) & "%' or shortName Like N'%" & Trim(astrSplit2tems2(j)) & "%') "
           StrWhere = StrWhere + "  and NOT (ItemName IS NULL) and NOT (shortName IS NULL) and  NOT (ItemCode IS NULL)"
         Next j
         End If
    sql = sql & StrWhere
        If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql + " Order BY ItemName "
    Else
        sql = sql + " Order BY ItemNamee "
    End If


   End If
   
        fill_combo DCboItemsName, sql
        
End Sub
Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId
    End If
End Sub

Private Sub TxtTransSerial_KeyDown(KeyCode As Integer, _
                                   Shift As Integer)
    Dim StrSearch As String
    Dim VarBookMark As Variant
    Dim Msg As String

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyReturn Then
            If Trim$(TxtTransSerial.Text) <> "" Then
                StrSearch = Trim$(TxtTransSerial.Text)

                If Not (rs.BOF Or rs.EOF) Then
                    If rs.EditMode = adEditNone Then
                        VarBookMark = rs.Bookmark
                        rs.find "Transaction_Serial='" & StrSearch & "'", , adSearchForward, adBookmarkFirst

                        If Not (rs.BOF Or rs.EOF) Then
                            Me.Retrive rs("Transaction_ID").value
                        Else
                            rs.Bookmark = VarBookMark
                            Msg = "Â–Â «·›« Ê—… €Ì— „ÊÃÊœ…...!!!"
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        End If
                    End If
                End If
            End If
        End If
    End If

End Sub

Private Sub TxtValueAdded_Change()
RelinVatGrid
End Sub

Private Sub VatGrid_Click()
RelinVatGrid
End Sub

Sub RelinVatGrid() 'vatvatvatmat
Dim i As Integer
Dim SmValu As Double
Dim k As Integer



If SystemOptions.PriceWithVAT = True Then: GoTo xx: Exit Sub

SmValu = 0
If FG.ColIndex("Vat") = -1 Then Exit Sub

With VatGrid
For i = 1 To .Rows - 1
If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
For k = FG.FixedRows To FG.Rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
FG.TextMatrix(k, FG.ColIndex("Vat")) = val(.TextMatrix(i, .ColIndex("Vat")))
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = val(.TextMatrix(i, .ColIndex("Vatyo")))
End If
Next k

SmValu = SmValu + val(.TextMatrix(i, .ColIndex("Vat")))
Else
For k = FG.FixedRows To FG.Rows - 1
If k = i And val(FG.TextMatrix(k, FG.ColIndex("Code"))) = val(.TextMatrix(i, .ColIndex("ItemID"))) And val(FG.TextMatrix(k, FG.ColIndex("Valu"))) = val(.TextMatrix(i, .ColIndex("Valu"))) Then
FG.TextMatrix(k, FG.ColIndex("Vat")) = 0
FG.TextMatrix(k, FG.ColIndex("Vatyo")) = 0
End If
Next k
End If

Next i
End With
TxtValueAdded.Text = Format(SmValu, ".##")
LblValueAdded.Caption = Format(SmValu, ".##")
xx:
Me.LblTotal.Caption = val(LblTotalAll.Caption) + val(TxtValueAdded.Text) - val(LblDiscountsTotal.Caption)
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VSFlexGrid1
If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then
Cancel = True
Exit Sub
End If
Select Case .ColKey(Col)
Case "TransPayedValue"
If .Cell(flexcpChecked, Row, .ColIndex("payed")) = flexChecked Then
Cancel = False
Else
End If

Case "NoteSerial1"
Cancel = True
Case "too"
Cancel = True
Case "NoteDate"
Cancel = True
Case "branch_name"
Cancel = True
Case "Note_Value"
Cancel = True
Case "PayedValue"
Cancel = True
Case "RemainingValue"
Cancel = True
Case "NetValue"
Cancel = True

End Select
End With
Cancel = True
End Sub


Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "payed"
If .Cell(flexcpChecked, Row, .ColIndex("payed")) = flexChecked Then
.TextMatrix(Row, .ColIndex("TransPayedValue")) = .TextMatrix(Row, .ColIndex("RemainingValue"))
Else
.TextMatrix(Row, .ColIndex("TransPayedValue")) = 0
End If
End Select
End With
RelineBuy
End Sub
Function AutoCalculate() As Boolean
Dim i As Integer
Dim NetValu As Double
Dim TempValu As Double
Dim RemainValu As Double
NetValu = val(LblTotal.Caption)
With VSFlexGrid1
For i = 1 To .Rows - 1
RemainValu = val(.TextMatrix(i, .ColIndex("RemainingValue")))
If NetValu >= RemainValu Then
TempValu = RemainValu
NetValu = NetValu - TempValu
Else
TempValu = NetValu
NetValu = 0
End If
If TempValu > 0 Then
  .TextMatrix(i, .ColIndex("TransPayedValue")) = TempValu
  .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
   End If
Next i
End With
If NetValu <> 0 Then
AutoCalculate = False
Else
AutoCalculate = True
End If
End Function
Sub DeleteBillBuy()
Dim i As Integer
Dim StrSQL As String
With VSFlexGrid1
 For i = .FixedRows To .Rows - 1
 If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
      StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
     End If
     Next i
 End With
End Sub

Public Sub XPBtnMove_Click(Index As Integer)

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
DisplayRec:
         Me.TxtModFlg.Text = ""
        Dim StrSQL As String
     StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=9 "
     
StrSQL = StrSQL & "  AND      BranchId in(" & Current_branchSql & ")"
            If SystemOptions.usertype <> UserAdminAll Then
           '     StrSQL = StrSQL & " AND   BranchId=" & Current_branch
            End If


     If SystemOptions.usertype <> UserAdminAll Then
 
          If SystemOptions.FixedCustomer = 1 Then
            StrSQL = StrSQL & " and  UserID = " & user_id
             End If
  
  
        Me.dcBranch.Enabled = True
      
      
    End If
    
            If SystemOptions.SortInvoiceByEntry Then
                StrSQL = StrSQL & " Order by Transaction_ID"
            Else
                StrSQL = StrSQL & " Order by noteserial1"
            End If
                
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


            If Not (rs.EOF Or rs.BOF) Then
                rs.MoveLast
            End If
        Me.TxtModFlg.Text = "R"
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
Exit Sub
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" And Not (Me.ActiveControl Is TxtTransSerial) Then
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
        If Cmd(3).Enabled = False Then Exit Sub
        Cmd_Click (3)
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
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 1

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                FG.SetFocus
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

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
lbl(97).Caption = "Smart Search"
CmdDele.Caption = "Delete"
lbl(36).Caption = "No.JL"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    Cmd(8).Caption = "Show Recive Vchr."
    Me.XPTab301.TabCaption(0) = "Items"
    ''/////////
    lbl(40).Caption = "Remarks"
    Cmd(9).Caption = "Print 2"
    lbl(39).Caption = "VAT No"
    lbl(65).Caption = "Curr"
    lbl(38).Caption = "VAT"
    lbl(104).Caption = "Total"
    Me.XPTab301.TabCaption(2) = "VAT"
    Label22.Caption = "Data of VAT"
    ChecVAT.RightToLeft = False
    ChecVAT.Caption = "Select All"
With VatGrid
.TextMatrix(0, .ColIndex("Valu")) = "Item Value"
.TextMatrix(0, .ColIndex("index")) = "Serial"
.TextMatrix(0, .ColIndex("select")) = "Select"
.TextMatrix(0, .ColIndex("Code")) = "Item Code"
.TextMatrix(0, .ColIndex("Name")) = "Item Name"
.TextMatrix(0, .ColIndex("Vatyo")) = "Percentage"
.TextMatrix(0, .ColIndex("Vat")) = "Value"
End With
lbl(41).Caption = "Data of Sales Invoices"
Command10.Caption = "Cancel"
Label27.Caption = "Total"
        With VSFlexGrid1

.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("InstalValue")) = "Installment Value"
.TextMatrix(0, .ColIndex("haveqest")) = "Have Installments"
.TextMatrix(0, .ColIndex("payed")) = "Select"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "Bill No"
.TextMatrix(0, .ColIndex("too")) = "Bill Supplier"
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("branch_name")) = "Branch"
.TextMatrix(0, .ColIndex("Note_Value")) = "Original value"
.TextMatrix(0, .ColIndex("PayedValue")) = "Payed Value"
.TextMatrix(0, .ColIndex("RemainingValue")) = "Remaining"
.TextMatrix(0, .ColIndex("TransPayedValue")) = "Payed Trans"
.TextMatrix(0, .ColIndex("NetValue")) = "Net value"
.TextMatrix(0, .ColIndex("Show")) = "Show"
.TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
End With
    ''//////////
    lbl(23).Caption = "Manual Recive Vchr."
    lbl(25).Caption = "Sales Person"
     lbl(35).Caption = "Cash Customer"
    Me.XPTab301.TabCaption(1) = "Notes"
    Label4.Caption = "Doc Type"
    Frame3.Caption = "Ge Data"
    Cmd(10).Caption = "Print Ge"
    opt(0).Caption = "Returned"
    opt(1).Caption = "Changed"
    lbl(10).Caption = "Date"
    lbl(24).Caption = "Type Dis"
    CmdOpenTrans.Caption = "View"
    lbl(34).Caption = "Value"
    lbl(84).Caption = "Phone"
    CmdSearchTrans.Caption = "Search"
lbl(37).Caption = "ManualNO"
    lbl(20).Caption = "Payment Method"
    XPChkPayType(0).Caption = "Cash"
    XPChkPayType(1).Caption = "Credit"
    XPChkPayType(2).Caption = "Cheque"
    Label3.Caption = "Branch"
    Label1.Caption = "Box"

    lbl(13).Caption = "Value"
    lbl(15).Caption = "Value"
    lbl(16).Caption = "Value"
    lbl(18).Caption = "Cheque#"

    lbl(12).Caption = "Index"
    lbl(14).Caption = "Index"
    lbl(22).Caption = "Box"
    lbl(21).Caption = " date"

    lbl(19).Caption = " Cheque date"

    lbl(17).Caption = "Bank"

    Me.Caption = "Sales returns"
    C1Elastic6.Caption = Me.Caption

    lbl(5).Caption = "ID"
    lbl(6).Caption = "Invoice Date"
    lbl(7).Caption = "Customer Name"
    lbl(8).Caption = "Store "

    lbl(9).Caption = "Payment Type"
    lbl(4).Caption = "Invoice#"
    Ele(2).Caption = "Return Type"
 
    lbl(3).Caption = " Total:"
    lbl(32).Caption = "Disc"
    lbl(33).Caption = " Net:"

    lbl(1).Caption = " By:"
    lbl(0).Caption = "Rec. Count:"

    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = " Case"
    lbl(28).Caption = " Serial"
    lbl(27).Caption = "QTY"
    lbl(26).Caption = "Price"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
  With FG
  .TextMatrix(0, .ColIndex("Select")) = "Select"
  .TextMatrix(0, .ColIndex("order_no")) = "No.Shipment"
  .TextMatrix(0, .ColIndex("OrderArrivalDate")) = "Date.Shipment"
  End With
    With Grid
  .TextMatrix(0, .ColIndex("PaymentName")) = "  Payment"
  .TextMatrix(0, .ColIndex("Value")) = "Value"
  .TextMatrix(0, .ColIndex("CardNo")) = "Card No"
  End With
End Sub

Private Sub Form_Load()
    Dim RsClients As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset
    Dim Dcombos As ClsDataCombos

    On Error GoTo ErrTrap
    If SystemOptions.AllowEditVaTManulay = True Then
txtManulaVat.Enabled = True
txtManulaVat.Visible = True
Else
txtManulaVat.Enabled = False
txtManulaVat.Text = 0
txtManulaVat.Visible = False
End If




    ScreenNameArabic = "„—œÊœ«  «·„»Ì⁄«  "
    ScreenNameEnglish = " Sales Return"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1", 220
  If SystemOptions.AllowReturnFIFO = True Then
  XPTab301.TabVisible(3) = True
  Else
  XPTab301.TabVisible(3) = False
  End If
 If True = True Then
 XPTab301.TabVisible(2) = True
 Else
 XPTab301.TabVisible(2) = False
 End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Set NewGrid.Grid = FG
    NewGrid.GridTrans = ReturnSalling
Set NewGrid.txtManulaVat = Me.txtManulaVat
     Set NewGrid.TxtValueAdded = TxtValueAdded
     Set NewGrid.VatGrid = Me.VatGrid
    Set NewGrid.TxtInvID = XPTxtBillID
    Set NewGrid.TxtModFlag = TxtModFlg
    Set NewGrid.txtTotal = XPTxtSum
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
      Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
  Set NewGrid.LBLGross = LBLGross
  
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.LblTotalQty = Me.LblTotalQty
     'Set NewGrid.LblTotal = Me.LblTota
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.txtManulaVat = Me.txtManulaVat
    
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.StoreName = Me.DCboStoreName
    Set NewGrid.TxtShortName = Me.TxtShortName
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.GrdTBar = Me.TBar
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    Set NewGrid.LblTotalAll = Me.LblTotalAll
  Set NewGrid.TxtItemCodeB = TxtItemCodeB
    Dim My_SQL As String
    My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    fill_combo dcBranch, My_SQL
'Dcombos.GetDocTypebyid (Me.DCDocTypes), 9, val(Me.Dcbranch.BoundText)

    If SystemOptions.usertype <> UserAdminAll Then
           If SystemOptions.BranchCanNotEdit = True Then
                             Me.dcBranch.Enabled = False
                             Else
                               Me.dcBranch.Enabled = True
                             End If
      '  XPDtbBill.Enabled = False
    End If

    Resize_Form Me, TransactionSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FG.WallPaper = BGround.Picture
    AddTip
    XPTab301.CurrTab = 0
    SetDtpickerDate XPDtbBill

    If SystemOptions.UserInterface = EnglishInterface Then

        With CboPayMentType
            .Clear
            .AddItem "Cash"
            .AddItem "Credit"
        End With


           With XPCboDiscountType
            .Clear
            .AddItem "No Discount"
            .AddItem "Discount Value"
            .AddItem "Discount %"
        End With
    Else

        With CboPayMentType
            .Clear
            .AddItem "‰ﬁœ«"
            .AddItem "¬Ã·"
        End With

       With XPCboDiscountType
            .Clear
            .AddItem "·«ÌÊÃœ Œ’„"
            .AddItem "Œ’„ »ﬁÌ„…"
            .AddItem "Œ’„ »‰”»…"
        End With
        
    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetBanks Me.DCboBankName
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetSalesRepData Me.DcboEmp
Dcombos.GetDocTypebyid Me.DCDocTypes, 9, 0
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboStoreName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboBankName
     StrSQL = " select id,code from currency"
    fill_combo Me.DcCurrency, StrSQL
    NewGrid.FillGrid

    If SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboRetrunType
            .Clear
            .AddItem "With bill "
            .AddItem "With out Bill"
        End With

    Else

        With Me.CboRetrunType
            .Clear
            .AddItem "≈— Ã«⁄ „ﬁÌœ(„— »ÿ »›« Ê—… »Ì⁄)"
            .AddItem "≈— Ã«⁄ €Ì— „ﬁÌœ(€Ì— „— »ÿ »›« Ê—… »Ì⁄)"
        End With

    End If


 If SystemOptions.HideCost = True Then
       FG.ColHidden(FG.ColIndex("ItemCostPrice")) = True
     
  End If
 
    StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=-9"
StrSQL = StrSQL & "  AND      BranchId  in(" & Current_branchSql & ")"
    If SystemOptions.usertype <> UserAdminAll Then
      '  StrSQL = StrSQL & " AND   BranchId=" & Current_branch
    End If

    StrSQL = StrSQL & " Order by Transaction_ID"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'XPBtnMove_Click 2
    'Me.TxtModFlg.Text = "R"
'DcboEmp.SetFocus
InvType = 9
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
    BillCustomer
        If CboPayMentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).Text = XPTxtSum.Text
            
            '  XPTxtValue(2).text = XPTxtSum.text
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).Text = XPTxtSum.Text
            '  XPTxtValue(2).text = XPTxtSum.text
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "", 220

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

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap
    Dim RsTest As ADODB.Recordset
    Dim StrSQL As String

    Select Case Me.TxtModFlg.Text

        Case "R"
            '   Me.Caption = "„—œÊœ«  «·„»Ì⁄« "
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
        
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            Me.DCboBankName.locked = True
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            XPTxtChqueNum.Enabled = False
            DCboBankName.Enabled = False
            XPTxtValue(2).Enabled = False
            XPDTPDueDate.Enabled = False
            FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            End If
        
            CboPayMentType.locked = True
            DtpDelayDate.Enabled = False
            Ele(4).Enabled = False
            CboRetrunType.locked = True
            TxtInvSerial.Enabled = False
        
        Case "N"
            '   Me.Caption = "„—œÊœ«  «·„»Ì⁄« ( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '      Me.XPBtnMove(0).Enabled = False
            '      Me.XPBtnMove(1).Enabled = False
            '      Me.XPBtnMove(2).Enabled = False
            '      Me.XPBtnMove(3).Enabled = False
        
            FG.Enabled = True
            FG.Rows = 2
             If SystemOptions.DateCanNotEdit = True Then
             Me.XPDtbBill.Enabled = False
             Else
             Me.XPDtbBill.Enabled = True
             End If
            XPDtbBill.value = Date
            Me.DBCboClientName.locked = False
            CboPayMentType.locked = False
            Me.DCboStoreName.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            CboPayMentType.ListIndex = 0
       
            DtpDelayDate.Enabled = True
            DtpDelayDate.value = Date
            XPDTPDueDate.value = Date
            Ele(4).Enabled = True
            CboItemCase.ListIndex = 0
            CboRetrunType.locked = False
            TxtInvSerial.Enabled = True

        Case "E"
            '   Me.Caption = "„—œÊœ«  «·„»Ì⁄« (  ⁄œÌ· )"
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
            If SystemOptions.DateCanNotEdit = True Then
             Me.XPDtbBill.Enabled = False
             Else
             Me.XPDtbBill.Enabled = True
             End If
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.DCboBankName.locked = False
            CboPayMentType.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPDTPDueDate.Enabled = True
            DtpDelayDate.Enabled = True

            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPayMentType.ListIndex = 0 Then
                CboPayMentType_Change
            End If

            FG.Editable = flexEDKbdMouse
       
            DBCboClientName_Change
            Ele(4).Enabled = True
            CboRetrunType.locked = False
            TxtInvSerial.Enabled = True
    End Select

    If SystemOptions.usertype <> UserAdminAll Then
 
       ' Me.dcBranch.Enabled = True
       ' XPDtbBill.Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub
Sub SaveValueAdded()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset

sql = "Select * from  TransactionValueAdded where 1=-1"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
With Me.VatGrid
For i = 1 To .Rows - 1
If val(.TextMatrix(i, .ColIndex("ItemID"))) <> 0 Then
rs2.AddNew
rs2("Transaction_ID").value = val(Me.XPTxtBillID.Text)
rs2("Transaction_Type").value = 9
rs2("ItemID").value = val(.TextMatrix(i, .ColIndex("ItemID")))
rs2("Vatyo").value = val(.TextMatrix(i, .ColIndex("Vatyo")))
rs2("Vat").value = val(.TextMatrix(i, .ColIndex("Vat")))
rs2("Valu").value = val(.TextMatrix(i, .ColIndex("Valu")))
If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
rs2("selectd").value = 1
Else
rs2("selectd").value = 0
End If

rs2.update
End If
Next i
End With
End Sub

Sub RetriveValueAdded()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.Rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 9) And (dbo.TransactionValueAdded.Transaction_ID = " & val(XPTxtBillID.Text) & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.Rows = .Rows + rs2.RecordCount
For i = 1 To .Rows - 1
 .TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(rs2("Vat").value), "", rs2("Vat").value)
.TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(rs2("Vatyo").value), "", rs2("Vatyo").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("select")) = IIf(IsNull(rs2("selectd").value), 0, rs2("selectd").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub
Sub RetriveValueAddedData(Optional Transaction_ID As Double)
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VatGrid.Clear flexClearScrollable, flexClearEverything
    VatGrid.Rows = 1
sql = " SELECT     dbo.TransactionValueAdded.Transaction_Type, dbo.TransactionValueAdded.Transaction_ID, dbo.TransactionValueAdded.Vat, dbo.TransactionValueAdded.Vatyo,"
sql = sql & " dbo.TransactionValueAdded.ItemID , dbo.TblItems.itemname, dbo.TblItems.Fullcode, dbo.TblItems.ItemNamee ,dbo.TransactionValueAdded.selectd ,dbo.TransactionValueAdded.Valu "
sql = sql & " FROM         dbo.TransactionValueAdded LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems ON dbo.TransactionValueAdded.ItemID = dbo.TblItems.ItemID"
sql = sql & " Where (dbo.TransactionValueAdded.Transaction_Type = 21) And (dbo.TransactionValueAdded.Transaction_ID = " & Transaction_ID & ")"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Me.VatGrid
rs2.MoveFirst
.Rows = .Rows + rs2.RecordCount
For i = 1 To .Rows - 1
 .TextMatrix(i, .ColIndex("index")) = i
.TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs2("ItemID").value), "", rs2("ItemID").value)
.TextMatrix(i, .ColIndex("Vat")) = IIf(IsNull(rs2("Vat").value), "", rs2("Vat").value)
.TextMatrix(i, .ColIndex("Vatyo")) = IIf(IsNull(rs2("Vatyo").value), "", rs2("Vatyo").value)
.TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("select")) = IIf(IsNull(rs2("selectd").value), 0, rs2("selectd").value)
.TextMatrix(i, .ColIndex("Valu")) = IIf(IsNull(rs2("Valu").value), 0, rs2("Valu").value)

If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemName").value), "", rs2("ItemName").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("ItemNamee").value), "", rs2("ItemNamee").value)
End If
rs2.MoveNext
Next i
End With
End If
RelinVatGrid
End Sub
Function RetriveQtyItem(Optional NoteSerial1 As String, Optional Item_ID As Double, Optional ColorID As Integer, Optional ClassId As Integer, Optional itemsize As Integer, Optional UnitID As Integer, Optional Transaction_ID As Double) As Double
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Integer
 '**************************************************************************
  StrSQL = "SELECT     dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.Item_ID, "
  StrSQL = StrSQL & "                     dbo.Transaction_Details.UnitId, SUM(dbo.Transaction_Details.ShowQty * isnull( dbo.Transaction_Details.FLgReturn,1)) AS smQty"
  StrSQL = StrSQL & "  FROM         dbo.Transaction_Details RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
  StrSQL = StrSQL & "   WHERE     (((dbo.Transactions.NoteSerial1 = N'" & NoteSerial1 & "') AND (dbo.Transactions.Transaction_Type = 21)) OR"
  StrSQL = StrSQL & "                  (   (dbo.Transactions.ReturnSerial = N'" & NoteSerial1 & "') AND (dbo.Transactions.Transaction_Type = 9)))"
  StrSQL = StrSQL & "  AND (dbo.Transaction_Details.Item_ID = " & Item_ID & ") AND (dbo.Transaction_Details.UnitId = " & UnitID & ") and(dbo.Transaction_Details.ColorID = " & ColorID & ") and(dbo.Transaction_Details.ClassId = " & ClassId & ")and(dbo.Transaction_Details.ItemSize = " & itemsize & ")"
StrSQL = StrSQL & "  AND   Transaction_Details.Transaction_ID <>" & val(Me.XPTxtBillID) & ""
  StrSQL = StrSQL & "  GROUP BY dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.ClassId, dbo.Transaction_Details.Item_ID,"
  StrSQL = StrSQL & "                     dbo.Transaction_Details.unitid"
    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsDetails.RecordCount > 0 Then
RetriveQtyItem = IIf(IsNull(RsDetails("smQty").value), 0, RsDetails("smQty").value)
Else
RetriveQtyItem = 0
End If
End Function
Function Retrive_Sales_invoice_data(Transaction_ID As Double)
    Dim StrSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim row_count As Integer
    Dim Num As Long
    
    Dim rs As ADODB.Recordset
    
 '**************************************************************************
 
        StrSQL = "Select * from transactions where  Transaction_ID=" & Transaction_ID
  

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Function
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    '    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        Me.DCboStoreName.BoundText = IIf(IsNull(rs("storeid").value), "", rs("storeid").value)
      Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
      Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
      
        
    If IsNull(rs("chkTaxExempt").value) Then
        Me.chkTaxExempt.value = vbUnchecked
    Else
        Me.chkTaxExempt.value = IIf(rs("chkTaxExempt").value = 0, vbUnchecked, vbChecked)
    End If
    
     
      
        Me.dcBranch.BoundText = IIf(IsNull(rs("Branchid").value), "", rs("Branchid").value)
Me.DcboBox.BoundText = IIf(IsNull(rs("bOXID").value), "", rs("bOXID").value)
    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
 
    XPTxtDiscountVal.Text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))
txtManulaVat.Text = IIf(IsNull(rs("txtManulaVat").value), 0, (rs("txtManulaVat").value))
txtManulaVat.Text = val(txtManulaVat.Text)

  End If
     '**************************************************************************
rs.Close
Set rs = Nothing
     
    
    
    
    
    
    
    StrSQL = "SELECT dbo.Transaction_Details.ItemSerial , TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.Rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.Rows - 1 'RsDetails.RecordCount
            FG.TextMatrix(Num, FG.ColIndex("itemtype")) = IIf(IsNull(RsDetails("itemtype")), "", Trim(RsDetails("itemtype").value))
            FG.TextMatrix(Num, FG.ColIndex("TotalDiscountPerLine")) = IIf(IsNull(RsDetails("TotalDiscountPerLine")), "", (RsDetails("TotalDiscountPerLine").value))
Debug.Print FG.TextMatrix(Num, FG.ColIndex("discountvalue"))
            FG.TextMatrix(Num, FG.ColIndex("Select")) = True
            FG.TextMatrix(Num, FG.ColIndex("EmpID4")) = IIf(IsNull(RsDetails("EmpID4")), "", (RsDetails("EmpID4").value))
            FG.TextMatrix(Num, FG.ColIndex("TypeVAT")) = IIf(IsNull(RsDetails("TypeVAT")), "", (RsDetails("TypeVAT").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(Num, FG.ColIndex("MaxQty")) = IIf(IsNull(RsDetails("MaxQty")), "", (RsDetails("MaxQty").value))
            FG.TextMatrix(Num, FG.ColIndex("MaxUnitID")) = IIf(IsNull(RsDetails("MaxUnitID")), "", (RsDetails("MaxUnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("MixNo")) = IIf(IsNull(RsDetails("MixNo")), "", (RsDetails("MixNo").value))
            
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '   FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("Valu")) = val(FG.TextMatrix(Num, FG.ColIndex("Price"))) * val(FG.TextMatrix(Num, FG.ColIndex("Count")))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
           FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", (RsDetails("ItemSerial").value))
           FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))

            End If

            
           
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            FG.TextMatrix(Num, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
            FG.TextMatrix(Num, FG.ColIndex("Area")) = IIf(IsNull(RsDetails("Area")), "", (RsDetails("Area").value))
            FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
            FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))
            
             FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
            FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

End Function

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim Num As Long
     On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.Text = "T"
    Screen.MousePointer = vbArrowHourglass
    TxtVATNO.Text = IIf(IsNull(rs("VATNO").value), "", (rs("VATNO").value))
    Me.DcCurrency.BoundText = IIf(IsNull(rs("Currency_id").value), 1, rs("Currency_id").value)
    txt_Currency_rate.Text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
    Me.Text1.Text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
  XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), -1, val(rs("Trans_DiscountType").value))
 TxtBillComment.Text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))
  txtManulaVat.Text = IIf(IsNull(rs("txtManulaVat").value), 0, (rs("txtManulaVat").value))
txtManulaVat.Text = val(txtManulaVat.Text)

   
    If IsNull(rs("chkTaxExempt").value) Then
        Me.chkTaxExempt.value = vbUnchecked
    Else
        Me.chkTaxExempt.value = IIf(rs("chkTaxExempt").value = 0, vbUnchecked, vbChecked)
    End If
    
    
    XPTxtDiscountVal.Text = IIf(IsNull(rs("Trans_Discount").value), "", (rs("Trans_Discount").value))

    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.Text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    TxtManualNo1.Text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))
    Me.oldtxtNoteSerial1.Text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(11).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TxtNoteID.Text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
'''/////////
 TxtValueAdded.Text = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))
  LblValueAdded.Caption = IIf(IsNull(rs("VAT").value), 0, (rs("VAT").value))


    XPTxtBillID.Text = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))
    FillGridWithDataActual val(XPTxtBillID.Text)
    TxtTransSerial.Text = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    FG.Clear flexClearScrollable, flexClearEverything
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)
    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If
    
    If Not IsNull(rs("ReturnID").value) Then
        Me.CboRetrunType.ListIndex = 0
        Me.TxtInvID.Text = rs("ReturnID").value
        
        
        Me.TxtInvSerial.Text = IIf(IsNull(rs("ReturnSerial").value), "", (rs("ReturnSerial").value))
        Me.txtInvDate.Text = IIf(IsNull(rs("SalesInvoiceDate").value), "", (rs("SalesInvoiceDate").value))


        If Not IsNull(rs("Returntype").value) Then
    
            If rs("Returntype").value = False Then
                opt(0).value = True
            Else
                opt(1).value = True
            End If
    
        Else
    
            opt(1).value = True
        End If
    
    Else
        Me.CboRetrunType.ListIndex = 1
        Me.TxtInvID.Text = ""
        Me.TxtInvSerial.Text = ""
        Me.txtInvDate.Text = ""

    End If
 ''//26 05 2015
    Me.TxtManualNO.Text = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))
    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    FG.Refresh
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
    StrSQL = StrSQL & " ORDER BY ID "

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.Text = ""

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
                    FG.TextMatrix(Num, FG.ColIndex("ParrtNoCode")) = IIf(IsNull(RsDetails("ParrtNoCode")), "", (RsDetails("ParrtNoCode").value))
FG.TextMatrix(Num, FG.ColIndex("ItemDetailedCode")) = IIf(IsNull(RsDetails("ItemDetailedCode")), "", (RsDetails("ItemDetailedCode").value))


 

            FG.TextMatrix(Num, FG.ColIndex("discountvalue")) = IIf(IsNull(RsDetails("discountvalue")), "", (RsDetails("discountvalue").value))
Debug.Print FG.TextMatrix(Num, FG.ColIndex("discountvalue"))
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))
            FG.TextMatrix(Num, FG.ColIndex("itemtype")) = IIf(IsNull(RsDetails("itemtype")), "", Trim(RsDetails("itemtype").value))
            FG.TextMatrix(Num, FG.ColIndex("TotalDiscountPerLine")) = IIf(IsNull(RsDetails("TotalDiscountPerLine")), "", (RsDetails("TotalDiscountPerLine").value))
            FG.TextMatrix(Num, FG.ColIndex("EmpID4")) = IIf(IsNull(RsDetails("EmpID4")), "", Trim(RsDetails("EmpID4").value))
            FG.TextMatrix(Num, FG.ColIndex("MaxQty")) = IIf(IsNull(RsDetails("MaxQty")), "", (RsDetails("MaxQty").value))
            FG.TextMatrix(Num, FG.ColIndex("MaxUnitID")) = IIf(IsNull(RsDetails("MaxUnitID")), "", (RsDetails("MaxUnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("MixNo")) = IIf(IsNull(RsDetails("MixNo")), "", (RsDetails("MixNo").value))
 'val(Fg.TextMatrix(RowNum, Fg.ColIndex("itemtype"))) <
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
            FG.TextMatrix(Num, FG.ColIndex("TypeVAT")) = IIf(IsNull(RsDetails("TypeVAT")), "", (RsDetails("TypeVAT").value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showprice")), "", (RsDetails("showprice").value))

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If
             FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
             FG.TextMatrix(Num, FG.ColIndex("Vatyo")) = IIf(IsNull(RsDetails("Vatyo")), "", (RsDetails("Vatyo").value))
             ''//////////////
            FG.TextMatrix(Num, FG.ColIndex("ParrtNoCode")) = IIf(IsNull(RsDetails("ParrtNoCode")), "", (RsDetails("ParrtNoCode").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCostPrice")) = IIf(IsNull(RsDetails("CostPrice")), "", (RsDetails("CostPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
           FG.TextMatrix(Num, FG.ColIndex("ItemsDetailsNewidea")) = IIf(IsNull(RsDetails("ItemsDetailsNewidea")), "", (RsDetails("ItemsDetailsNewidea").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
            FG.TextMatrix(Num, FG.ColIndex("NoCount")) = IIf(IsNull(RsDetails("NoCount")), "", (RsDetails("NoCount").value))
            FG.TextMatrix(Num, FG.ColIndex("Area")) = IIf(IsNull(RsDetails("Area")), "", (RsDetails("Area").value))
            FG.TextMatrix(Num, FG.ColIndex("Height")) = IIf(IsNull(RsDetails("Height")), "", (RsDetails("Height").value))
            FG.TextMatrix(Num, FG.ColIndex("Width")) = IIf(IsNull(RsDetails("Width")), "", (RsDetails("Width").value))

            RsDetails.MoveNext
            Debug.Print Num

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

    End If

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).Text = ""
    XPTxtValue(1).Text = ""
    XPTxtValue(2).Text = ""
    XPTxtSerial(0).Text = ""
    XPTxtSerial(1).Text = ""
    XPTxtChqueNum.Text = ""
    DCboBankName.BoundText = ""
    XPDTPDueDate.value = Date
    DtpDelayDate.value = Date
    StrSQL = "select * From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
    RsNotes.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                XPTxtValue(0).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
            
                XPTxtSerial(0).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", (RsNotes("BoxID").value))
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                XPTxtValue(1).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).Text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 13 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
                XPTxtValue(2).Text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtChqueNum.Text = IIf(IsNull(RsNotes("ChqueNum").value), "", Trim(RsNotes("ChqueNum").value))
                Me.DCboBankName.BoundText = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                XPDTPDueDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            RsNotes.MoveNext
        Next Num

    End If
    If Me.TxtModFlg.Text <> "E" And Me.TxtModFlg.Text <> "N" Then
   RetriveBillBuyData
   RetriveValueAdded
   RelinVatGrid
       
    End If

 
 '           NewGrid.Calculate 1, , , True
 '      DoEvents
 '      NewGrid.SentTypeVAT
       
    TxtFillData.Text = "F"
 
       
    Screen.MousePointer = vbDefault
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
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
                rs.find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst

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
 Function saveBillBuy()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
      If Me.TxtModFlg.Text = "E" Then
  '  StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtBillID.Text) & " and TransType=1"
  '  Cn.Execute StrSQL, , adExecuteNoRecords
    
    StrSQL = "Delete From TblBillBuyPayment2 Where     NoteID=" & general_noteid & " and TransType=1"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    End If

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If CboPayMentType.ListIndex = 0 Then Exit Function
    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
    TxtValueTemp.Text = val(LblTotal.Caption)
    For i = .FixedRows To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
            RsDetails.AddNew
            RsDetails("NoteID1").value = val(TxtNoteID.Text)  'val(XPTxtBillID.Text)
            RsDetails("TransType").value = 1
            RsDetails("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
            RsDetails("branch_no").value = val(.TextMatrix(i, .ColIndex("branch_no")))
            RsDetails("NoteSerial1").value = val(.TextMatrix(i, .ColIndex("NoteSerial1")))
            RsDetails("Note_Value").value = val(.TextMatrix(i, .ColIndex("Note_Value")))
            Note_Value1 = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            Diff = 0
            If val(TxtValueTemp.Text) > 0 Then
          If val(TxtValueTemp.Text) <= Note_Value1 Then
          Diff = val(TxtValueTemp.Text)
          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
          Else
          Diff = Note_Value1
          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
          End If
            End If
          ' .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("RemainingValue")))
            .TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
            
            RsDetails("PayedValue").value = val(.TextMatrix(i, .ColIndex("PayedValue")))
            
            RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
            RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
            If .TextMatrix(i, .ColIndex("DueDate")) <> "" And .TextMatrix(i, .ColIndex("DueDate")) <> " " Then
            RsDetails("DueDate").value = IIf((.TextMatrix(i, .ColIndex("DueDate"))) = "", Null, (.TextMatrix(i, .ColIndex("DueDate"))))
            Else
            RsDetails("DueDate").value = Null
            End If
            RsDetails("TransPayedValue").value = val(.TextMatrix(i, .ColIndex("TransPayedValue")))
           .TextMatrix(i, .ColIndex("NetValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("TransPayedValue")))
            RsDetails("NetValue").value = val(.TextMatrix(i, .ColIndex("NetValue")))
            RsDetails("RemainingValue").value = val(.TextMatrix(i, .ColIndex("RemainingValue")))
            RsDetails.update
                
            If val(val(.TextMatrix(i, .ColIndex("NetValue")))) = 0 Then
            StrSQL = "Update Transactions Set  TotalPayed=1 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(.TextMatrix(i, .ColIndex("NoteID"))) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
      End If
    Next i
End With
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With VSFlexGrid1
  '  For i = .FixedRows To .Rows - 1
        If CboPayMentType.ListIndex = 1 Then
            RsDetails.AddNew
            RsDetails("NoteID").value = general_noteid
            RsDetails("RecDate").value = XPDtbBill.value
            RsDetails("Serial").value = TxtNoteSerial1.Text
            RsDetails("TransType").value = 1
            RsDetails("Transaction_ID").value = val(Me.TxtInvID.Text)
            RsDetails("Note_Value").value = val(LblTotal.Caption)
            RsDetails("PayedValue").value = val(LblTotal.Caption) * -1
            RsDetails.update
        End If
  '  Next i
End With

End Function
Private Sub Del_TransAction()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTest As ADODB.Recordset
    Dim BegainTrans As Boolean

    On Error GoTo ErrTrap

    If XPTxtBillID.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtNoteSerial1.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then

            '«·√ﬁ”«ÿ «·„”œœ… ⁄·Ï «·›« Ê—…
            If XPTxtValue(1).Tag <> "" Then
                StrSQL = "select * From ReceiptQestForBill where NoteID=" & XPTxtValue(1).Tag
                Set RsTest = New ADODB.Recordset
                RsTest.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If Not (RsTest.EOF Or RsTest.BOF) Then
                    Msg = "·ﬁœ  „  Õ’Ì· »⁄÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï Â–Â «·›« Ê—…" & CHR(13)
                    Msg = Msg + "Ê·« Ì„ﬂ‰ Õ–› »Ì«‰« Â«" & CHR(13)
                    Msg = Msg + "≈–« ﬂ‰   —€» ›Ì Õ–› »Ì«‰«  Â–Â «·›« Ê—…" & CHR(13)
                    Msg = Msg + "ÌÃ» Õ–› ⁄„·Ì«  «· Õ’Ì· «·Œ«’… »Â«"
                    MsgBox Msg, vbOKOnly + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            End If

            If Not rs.RecordCount < 1 Then
                Cn.BeginTrans
                BegainTrans = True
                Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.Text) & ""
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
            
                Cn.Execute StrSQL, , adExecuteNoRecords
            
                StrSQL = "delete From Notes where noteid=" & val(TxtNoteID.Text)
    
                Cn.Execute StrSQL, , adExecuteNoRecords
                    DeleteBillBuy
              StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.XPTxtBillID.Text) & " and TransType=1"
              Cn.Execute StrSQL, , adExecuteNoRecords
        '      StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.XPTxtBillID.Text) & " and TransType=1"
        '      Cn.Execute StrSQL, , adExecuteNoRecords
                
         StrSQL = "Delete From TblBillBuyPayment2 Where     NoteID=" & val(Me.TxtNoteID.Text) & " and TransType=1"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
                DeleteTransactiomsVoucher val(Text1.Text)
                
                rs.delete
                CuurentLogdata ("D")
                Cn.CommitTrans
                BegainTrans = False
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                     VatGrid.Clear flexClearScrollable, flexClearEverything
                     VatGrid.Rows = 1
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
         VatGrid.Clear flexClearScrollable, flexClearEverything
           VatGrid.Rows = 1
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.description
    Msg = Msg & CHR(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "· ”ÃÌ· ⁄„·Ì… „—œÊœ«  „»Ì⁄«  ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·⁄„·Ì… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·≈÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·⁄„·Ì… «·Õ«·Ì…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… „—œÊœ«  „»Ì⁄« " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  „—œÊœ«  «·„»Ì⁄« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT     dbo.TblPaymentType.PaymentID, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.BankId, dbo.TblPaymentType.Accountsus, "
My_SQL = My_SQL & "  dbo.TblPaymentType.Accountcom, dbo.TblPaymentType.commision, dbo.TblPaymentType.PaymentNamee, dbo.BanksData.Account_Code AS bankAccount_Code"
My_SQL = My_SQL & " FROM         dbo.TblPaymentType LEFT OUTER JOIN"
My_SQL = My_SQL & " dbo.BanksData ON dbo.TblPaymentType.BankId = dbo.BanksData.BankID order by PaymentID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 2
            rs.MoveFirst
      If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(1, .ColIndex("PaymentName")) = " ‰ﬁœÌ"
               Else
               .TextMatrix(1, .ColIndex("PaymentName")) = " Cash"
               End If
               
                .TextMatrix(1, .ColIndex("PaymentID")) = 0
           
           
            For i = 2 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "", rs.Fields("PaymentNamee").value)
               End If
               
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), "", rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("BankId")) = IIf(IsNull(rs.Fields("BankId").value), "", rs.Fields("BankId").value)
            
            .TextMatrix(i, .ColIndex("Accountsus")) = IIf(IsNull(rs.Fields("Accountsus").value), "", rs.Fields("Accountsus").value)
            .TextMatrix(i, .ColIndex("Accountcom")) = IIf(IsNull(rs.Fields("Accountcom").value), "", rs.Fields("Accountcom").value)
            .TextMatrix(i, .ColIndex("commision")) = IIf(IsNull(rs.Fields("commision").value), "", rs.Fields("commision").value)
           .TextMatrix(i, .ColIndex("bankAccount_Code")) = IIf(IsNull(rs.Fields("bankAccount_Code").value), "", rs.Fields("bankAccount_Code").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub




Public Sub FillGridWithDataActual(Transaction_ID As Double)

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "  SELECT     TOP 100 PERCENT dbo.TblTransactionPayments.[value], dbo.TblTransactionPayments.Effect, dbo.TblTransactionPayments.Transaction_ID, "
My_SQL = My_SQL & "   dbo.TblPaymentType.PaymentID"
My_SQL = My_SQL & ", dbo.TblPaymentType.PaymentName , dbo.TblTransactionPayments.CardNo  FROM         dbo.TblTransactionPayments LEFT OUTER JOIN"
My_SQL = My_SQL & "   dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
My_SQL = My_SQL & " Where (dbo.TblTransactionPayments.Transaction_ID = " & Transaction_ID & ")"
My_SQL = My_SQL & " ORDER BY dbo.TblPaymentType.PaymentID"

     With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

End With
    
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 2
            rs.MoveFirst


            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
             
             
             If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentName").value), "‰ﬁœÌ", rs.Fields("PaymentName").value)
               Else
               .TextMatrix(i, .ColIndex("PaymentName")) = IIf(IsNull(rs.Fields("PaymentNamee").value), "Cash", rs.Fields("PaymentNamee").value)
               End If
               
                .TextMatrix(i, .ColIndex("PaymentID")) = IIf(IsNull(rs.Fields("PaymentID").value), 0, rs.Fields("PaymentID").value)
           
                .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs.Fields("Value").value), 0, rs.Fields("Value").value)
            
            .TextMatrix(i, .ColIndex("CardNo")) = IIf(IsNull(rs.Fields("CardNo").value), "", rs.Fields("CardNo").value)
       
                rs.MoveNext
            Next

            rs.Close
        End If

  '      .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub



Function CheckAccount() As Boolean
Dim StrTempAccountCode As String
Dim usedaccount As Integer
Dim Account_Code_dynamic As String
    CheckAccount = False
    'Dcombos.GetDocTypebyid Me.DCDocTypes, 21, val(Me.dcBranch.BoundText)

    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«»  «·„œÌ‰ «·Œ«’ »«·„—œÊœ«   ", vbCritical
            GoTo ErrTrap
        End If
               
    End If



    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·„—œÊœ«   «·„»Ì⁄« ", vbCritical
            GoTo ErrTrap
        End If
 
    End If


    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·”‰œ  «·«” ·«„", vbCritical
            GoTo ErrTrap
        End If
 
    End If
    
   If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ  «·«” ·«„", vbCritical
            GoTo ErrTrap
        End If
 
    End If
     
     
    
    
    
   If SystemOptions.DiscountSalesCreateVchr = True Then
     If val(Me.LblDiscountsTotal.Caption) > 0 Then
           
    If val(DCDocTypes.BoundText) > 0 Then
        getDocAccounts val(DCDocTypes.BoundText), , , StrTempAccountCode, , , , , usedaccount

        If StrTempAccountCode = "" And usedaccount = 1 Then
            MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·Œ«’ »«·Œ’„ «·„”„ÊÕ »Â ", vbCritical
            GoTo ErrTrap
        End If
               
    End If
           
           
           
           
           Account_Code_dynamic = get_account_code_branch(12, my_branch)
    
           If Account_Code_dynamic = "NO branch" Then
                       If SystemOptions.UserInterface = ArabicInterface Then
                           MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                       Else
                       MsgBox "Branch Not Created ", vbCritical
                       End If
               GoTo ErrTrap
           ElseIf Account_Code_dynamic = "NO account" Then
                               If SystemOptions.UserInterface = ArabicInterface Then
                                   MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·Œ’„ «·„”„ÊÕ »Â   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                               Else
                               MsgBox "Allowance Discount Not Deined in this Branch", vbCritical
                               End If
                   GoTo ErrTrap
    
          End If
            
    End If
End If

  
    CheckAccount = True
    Exit Function
ErrTrap:

    CheckAccount = False
End Function
Private Sub SaveData()
Dim usedaccount As Integer
    Dim Msg As String
    Dim RowNum As Integer
    Dim RSTransDetails As ADODB.Recordset
  Dim RSTransDetails1 As ADODB.Recordset
        '    rs("Transaction_NetValue").value = val(lblInstComm.Caption) + val(LblTotal.Caption) + val(Me.TxtValueAdded.Text)
    
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp  As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RsRepeat As ADODB.Recordset
    Dim RsDetalis As ADODB.Recordset
    Dim StrSQL As String
    Dim StrSqlDel As String
    Dim note_id As Long
    Dim BeginTrans As Boolean
         Dim TotalDiscountPerLine As Variant
    Dim TotalBillDiscount As Double
        Dim ItemsGoodsTotalsnew As Variant
        Dim ItemsServiceTotalsnew As Variant
        TotalDiscountPerLine = 0
        TotalBillDiscount = 0
        ItemsGoodsTotalsnew = 0
        ItemsServiceTotalsnew = 0
    '
  '  On Error GoTo ErrTrap
    Screen.MousePointer = vbArrowHourglass

    If Me.TxtModFlg.Text <> "R" Then
        If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "≈–« ﬂ«‰ Â‰«ﬂ Œ’„ ⁄·Ï «·›« Ê—… " & CHR(13)
                Msg = Msg + "ÌÃ»  ÕœÌœ ﬁÌ„… Â–« «·Œ’„ " & CHR(13)
                Msg = Msg + "√Ê √Œ Ì«— ·« ÌÊÃœ Œ’„ "
            Else
                Msg = Msg + " Must Enter Discount Value " & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPCboDiscountType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
        If DBCboClientName.Text = "" Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If DCboStoreName.Text = "" Then
            Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DCboStoreName.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If CboPayMentType.ListIndex = -1 Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CboPayMentType.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.XPChkPayType(0).value = vbChecked Then
            If Me.DcboBox.BoundText = "" Then
                Msg = "ÌÃ»  ÕœÌœ «·Œ“‰…...!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).Text), XPDtbBill.value) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        If XPChkPayType(2).value = vbChecked Then
            If DCboBankName.BoundText = "" Then
                Screen.MousePointer = vbDefault
                MsgBox "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            If Trim(Me.XPTxtChqueNum.Text) = "" Then
                Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!!"
                Screen.MousePointer = vbDefault
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            If Check_CheckNum(Me.XPTxtChqueNum.Text, val(Me.XPTxtBillID.Text), Me.TxtModFlg.Text, 0) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If

        If Me.Ele(2).Visible = True Then
            If Me.CboRetrunType.ListIndex = -1 Then
                Msg = "»—Ã«¡ ≈Œ Ì«— ‰Ê⁄ «·√— Ã«⁄.."
                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                CboRetrunType.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            ElseIf Me.CboRetrunType.ListIndex = 0 Then

                If Trim(Me.TxtInvSerial.Text) = "" Then
                    Msg = "›Ï Õ«·… «·√— Ã«⁄ «·„ﬁÌœ »›« Ê—… »Ì⁄ "
                    Msg = Msg & CHR(13) & "ÌÃ» ﬂ «»… —ﬁ„ ›« Ê—… «·»Ì⁄"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    '                TxtInvSerial.SetFocus
                    Screen.MousePointer = vbDefault
                    Exit Sub
                ElseIf CheckInvData = False Then
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If

        '----------------------------------------------------------------------------
        'Check the Items Grid
        Me.XPTab301.CurrTab = 0

        If NewGrid.CheckDataEntered = False Then
            Exit Sub
        End If
        If CheckAccount = False Then
        Exit Sub
        End If

        '----------------------------------------------------------------------------
        '    If NewGrid.GetItemsCostTotal = 0 Then
        '        Msg = "«·»—‰«„Ã €Ì— ﬁ«œ— ⁄·Ï Õ”«»  ﬂ·›… «·√’‰«› «·„ÊÃÊœ… ›Ï ⁄„·Ì… «·„—œÊœ« "
        '        Msg = Msg & Chr(13) & "»—Ã«¡ „—«Ã⁄… √”⁄«—  ﬂ·›… «·√’‰«›"
        '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
   
        ' „—«Ã⁄Â «”⁄«— «· ﬂ·›…
        
        
        Cn.BeginTrans
        BeginTrans = True
        Screen.MousePointer = vbArrowHourglass
        
        DeleteTransactiomsVoucher val(Text1.Text)
        Dim UnitID As Long
        Dim MsgBoxResult As Integer
        Dim DblItemCostPrice  As Double
        If SystemOptions.AllowReturnWithoutCost = True Then
                       For RowNum = 1 To FG.Rows - 1
                       If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                    UnitID = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", 0, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                    FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.Text), UnitID, val(Me.DCboStoreName.BoundText)))
                    End If
                    
                               If val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) = 0 Then
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Price")))
                        End If
                        
                    Next RowNum
     End If
If SystemOptions.AllowReturnWithoutCost = False Then
        For RowNum = 1 To FG.Rows - 1
     
             If val(FG.TextMatrix(RowNum, FG.ColIndex("itemtype"))) <> 0 Then
                       FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Price")))
                End If
     
            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                
                If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
                    If val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))) = 0 Then
                        MsgBox "«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ "
                                              
                    GoTo ErrTrap
                    End If
                                 
                Else '€Ì— „ﬁÌœ »›« Ê—…
                    UnitID = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))

                    If val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.Text), UnitID, val(DCboStoreName.BoundText))) = 0 Then
                        'If Val(ModItemCostPrice.GetCostItemPrice(Fg.TextMatrix(RowNum, Fg.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod)) = 0 Then
                        MsgBoxResult = MsgBox("«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â —»„« ·⁄œ„ ÊÃÊœ ﬂ„Ì… Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ " & CHR(13) & "Â·  —Ìœ Õ”«»  ﬂ·› … ⁄·Ï «”«” «Œ— ”‰œ ’—› «‰ ÊÃœ ‰⁄„ «Ê ·« ", vbYesNo)

                        If MsgBoxResult = vbYes Then
                            FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = getLastCostPriceForItems(FG.TextMatrix(RowNum, FG.ColIndex("Code")), UnitID)
                        Else
                            MsgBoxResult = MsgBox("«·’‰›   " & FG.TextMatrix(RowNum, FG.ColIndex("Name")) & " €Ì— „Õœœ ”⁄—  ﬂ·› Â —»„« ·⁄œ„ ÊÃÊœ ﬂ„Ì… Ê·–·ﬂ ·« Ì„ﬂ‰ « „«„ ⁄„·ÌÂ «·«—Ã«⁄ " & CHR(13) & "Â·  —Ìœ Õ”«»  ﬂ·› … ⁄·Ï «”«” «‰ ÌﬂÊ‰ ‰›” ”⁄— «·„—œÊœ«  ‰⁄„ / ·« «ﬁÊ„ »√œŒ«· ”⁄— ÌœÊÌ ", vbYesNo)

                            If MsgBoxResult = vbYes Then
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = FG.TextMatrix(RowNum, FG.ColIndex("Price"))
                            Else
                                DblItemCostPrice = InputBox("«œŒ· «·”⁄— ··’‰›" & FG.TextMatrix(RowNum, FG.ColIndex("Name")))
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(DblItemCostPrice)
                            End If
                                                                             
                        End If
                                                    
                    Else
                        FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(ModItemCostPrice.GetCostItemPrice(FG.TextMatrix(RowNum, FG.ColIndex("Code")), 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbBill.value, val(Text1.Text), UnitID, val(Me.DCboStoreName.BoundText)))
                        '   Exit Sub
                        
                        If FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = 0 Then
                                FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = val(FG.TextMatrix(RowNum, FG.ColIndex("Price")))
                        End If

                    End If
                                            
                End If
                      
            End If

        Next RowNum
      End If
        If CheckRetrunInv = False Then
            Screen.MousePointer = vbDefault
             GoTo ErrTrap
        End If

        '  If NewGrid.Calculate(1, , True) = False Then
        '      Screen.MousePointer = vbDefault
        '      Exit Sub
        '  End If
    
        CurrentVoucherNo = ""
        CurrentVoucherSerialNo = ""

        'Create big notes
     my_branch = val(Me.dcBranch.BoundText)
        If TxtNoteSerial.Text = "" Then
            If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            Else
                       
                If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                Else
                    TxtNoteSerial.Text = Notes_coding(val(my_branch), XPDtbBill.value)
                End If
            End If
        End If
        
        Dim TxtNoteSerial1str As String
        
        If TxtNoteSerial1.Text = "" Then
        TxtNoteSerial1str = Voucher_coding(val(my_branch), XPDtbBill.value, 14, 220, , 9, , val(DCboStoreName.BoundText), , , , val(DCboUserName.BoundText))
            If TxtNoteSerial1str = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ „—œÊœ«  „»Ì⁄«   ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
            Else
                       
                If TxtNoteSerial1str = "" Then
                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„  ”‰œ «·«—Õ«⁄ ÌœÊÌ«  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                Else
                    TxtNoteSerial1.Text = TxtNoteSerial1str
                End If
            End If
        End If
     
        If Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20, , val(DCboStoreName.BoundText), , , , val(DCboUserName.BoundText)) = "" Then
                                
            If Trim$(TxtManualNo1) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·«” ·«„ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            
            Else
                TxtNoteSerial1V = TxtManualNo1
            End If
            
        End If
                    
        Set RsNotesGeneral = New ADODB.Recordset
'        RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
      StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
     RsNotesGeneral.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
        If Me.TxtModFlg.Text = "N" Then
            Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
            TxtNoteSerial1V = ""
        Else
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        
            StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
        
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            StrSQL = "delete From Notes where noteid=" & val(TxtNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
            CurrentVoucherNo = GetVoucherGLNO(val(Text1.Text), CurrentVoucherSerialNo)

            'DeleteTransactiomsVoucher val(Text1.text)
        
            general_noteid = val(TxtNoteID.Text)
            
                     StrSQL = "Delete From TblTransactionPayments Where Transaction_ID=" & val(Me.XPTxtBillID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If

        RsNotesGeneral.AddNew
        RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        general_noteid = RsNotesGeneral("NoteID").value
        TxtNoteID.Text = general_noteid
        ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
        RsNotesGeneral("NoteDate").value = XPDtbBill.value
        RsNotesGeneral("NoteType").value = 220
        RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
        RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.Text) = "", Null, Trim(Me.TxtNoteSerial.Text))
        RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
RsNotesGeneral("remark").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))

        RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        
        RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '??? ?????
        RsNotesGeneral("numbering_type1").value = sand_numbering_type(14) '  ?????? ???
        RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
        RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
        RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
'        RsNotesGeneral("ReturnInvoiceNO").value = (Me.TxtInvSerial.Text)
        
        'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
        RsNotesGeneral.update

        '---------------------------------
    
        Set RSTransDetails = New ADODB.Recordset
'        RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
 
        Set RsNotes = New ADODB.Recordset
'        RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
         StrSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText



        If Me.TxtModFlg.Text = "N" Then
            XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=9"))
            rs.AddNew
            rs("Transaction_ID").value = val(XPTxtBillID.Text)
            
        ElseIf Me.TxtModFlg.Text = "E" Then

            If rs("Transaction_ID").value <> val(XPTxtBillID.Text) Then
                rs.find "Transaction_ID=" & val(XPTxtBillID.Text) & "", , adSearchForward, 1
            End If
        End If

        rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
    
      If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If
    rs("txtManulaVat").value = val(txtManulaVat.Text)
    
        rs("VATNO").value = IIf(Trim(Me.TxtVATNO.Text) = "", Null, Trim(Me.TxtVATNO.Text))
        rs("Trans_Discount").value = IIf(XPTxtDiscountVal.Text = "", Null, val(XPTxtDiscountVal.Text))
        rs("VAT").value = val(TxtValueAdded.Text)
        rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
        rs("PPointID").value = val(txtPPointID)
        rs("ManualNo1").value = IIf(TxtManualNo1.Text = "", Null, val(TxtManualNo1.Text))
        rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.Text) = "", Null, Trim(Me.TxtNoteSerial.Text))
        rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.Text) = "", Null, Trim(Me.TxtNoteSerial1.Text))
        rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rs("NoteId").value = val(TxtNoteID.Text)
        rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.Text) = "", "", Trim(Me.TxtTransSerial.Text))
        rs("Transaction_Date").value = XPDtbBill.value
        rs("Transaction_Type").value = 9
        rs("UserID").value = user_id
        rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, val(DCboStoreName.BoundText))
        rs("BoxID").value = IIf(DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
        rs("TransactionComment").value = IIf(Trim$(TxtBillComment.Text) = "", Null, Trim$(TxtBillComment.Text))


            
        If chkTaxExempt.value = vbChecked Then
            rs("chkTaxExempt").value = 1
        Else
            rs("chkTaxExempt").value = 0
        End If

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

        If opt(0).value = True Then
            rs("ReturnType").value = 0
        Else
            rs("ReturnType").value = 1
        End If
   
        If CboPayMentType.ListIndex = -1 Then
    
            rs("PaymentType").value = 0
        Else
            rs("PaymentType").value = val(CboPayMentType.ListIndex)
        End If
     rs("Currency_id").value = IIf(DcCurrency.BoundText = "", Null, val(DcCurrency.BoundText))
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.Text), 1, txt_Currency_rate.Text)
        If Me.CboRetrunType.ListIndex = 0 Then
            rs("ReturnID").value = val(Me.TxtInvID.Text)
            rs("ReturnSerial").value = Me.TxtInvSerial.Text
            rs("SalesInvoiceDate").value = IIf(IsDate(Me.txtInvDate.Text), Me.txtInvDate.Text, Null)
        
        Else
            rs("ReturnID").value = Null
            rs("ReturnSerial").value = Null
            rs("SalesInvoiceDate").value = Null
        End If
   ''//26 05 2015
rs("ManualNO").value = IIf(Me.TxtManualNO.Text = "", Null, TxtManualNO.Text)
rs("Transaction_NetValue").value = val(LblTotal.Caption) ' + val(Me.TxtValueAdded.Text)

        rs.update
    
        If Me.TxtModFlg.Text = "E" Then
            Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.XPTxtBillID.Text) & ""
            StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
            StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSqlDel, , adExecuteNoRecords
       
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(rs("Transaction_ID").value)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
        End If
 'Dim TotalDiscountPerLine As Variant

        
            If Me.XPCboDiscountType.ListIndex = 1 Then
                     TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.Text <> "" Then
                 
                    TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text)) * val(LBLGross.Caption) / 100
                    
                    
                Else
                    TotalBillDiscount = 0
                End If
            End If
 
 
        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                                      RSTransDetails("ParrtNoCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))))
  RSTransDetails("ItemDetailedCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ItemDetailedCode"))))
  
  
                   RSTransDetails("ParrtNoCode").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ParrtNoCode"))))
                RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
                RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
                RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
                RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
                RSTransDetails("Transaction_ID").value = val(XPTxtBillID.Text)
                RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                Dim cnt As Double
                cnt = FG.TextMatrix(RowNum, FG.ColIndex("Count"))
                
                If LblTotalAll.Caption > 0 Then
     '  TotalDiscountPerLine = (RSTransDetails("SHOWprice") * RSTransDetails("SHOWQTY")) / LblTotalAll.Caption * (TotalBillDiscount)
        '   TotalDiscountPerLine = Fg.TextMatrix(RowNum, Fg.ColIndex("Valu")) / LblTotalAll.Caption * (TotalBillDiscount)
           If val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))) > 0 Then
           
        '   TotalDiscountPerLine = Fg.TextMatrix(RowNum, Fg.ColIndex("Valu")) / (LblFinal + (TotalBillDiscount)) * TotalBillDiscount
           TotalDiscountPerLine = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / (LBLGross) * TotalBillDiscount
           
         TotalDiscountPerLine = Round(TotalDiscountPerLine, 20)
           Else
           TotalDiscountPerLine = 0
           End If
         If val(FG.TextMatrix(RowNum, FG.ColIndex("itemtype"))) = 1 Then
                                                                                
         ItemsServiceTotalsnew = ItemsServiceTotalsnew + TotalDiscountPerLine + val(FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")))
         Else
         ItemsGoodsTotalsnew = ItemsGoodsTotalsnew + TotalDiscountPerLine + val(FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")))
         End If
 Else
 TotalDiscountPerLine = 0
 End If
RSTransDetails("TotalDiscountPerLine") = Round(TotalDiscountPerLine, 20)
'RSTransDetails("TotalDiscountPerLine") = val(Fg.TextMatrix(RowNum, Fg.ColIndex("TotalDiscountPerLine")))

                '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) = ""), Null, Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
                
                If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                    StrSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If RsTemp("HaveSerial").value = True Then
                            RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("Serial")))
                        End If
                    End If

                    RsTemp.Close
                End If
'''///////////
                RSTransDetails("EmpID4").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("EmpID4")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("EmpID4"))))
                RSTransDetails("MaxQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MaxQty")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("MaxQty"))))
                RSTransDetails("MixNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MixNo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("MixNo"))))
                RSTransDetails("MaxUnitID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("MaxUnitID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("MaxUnitID"))))
                RSTransDetails("TypeVAT").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("TypeVAT"))))
                RSTransDetails("Vat").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vat")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vat"))))
                RSTransDetails("Vatyo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Vatyo")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Vatyo"))))
                'Me.LblCostPrice.Caption = ModItemCostPrice.GetCostItemPrice(Val(Me.XPTxtID.text), 2)
                RSTransDetails("sallReturnPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
                RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
                RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
                RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
                'RSTransDetails("CostPrice").Value = Val(Fg.TextMatrix(RowNum, Fg.ColIndex("ItemCostPrice")))
                RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
                RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
                RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
                RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
                RSTransDetails("ItemsDetailsNewidea").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("ItemsDetailsNewidea")))
                RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
                ' RSTransDetails("price").value = ModItemCostPrice.GetCostItemPrice(Val(RSTransDetails("Item_ID").value), 2) * RSTransDetails("ShowQty").value
                RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
                RSTransDetails("FLgReturn").value = -1
                'RSTransDetails("CostPrice").value = RSTransDetails("PRICE").value * RSTransDetails("quantity").value
        
                If CboRetrunType.ListIndex = 0 Then '„ﬁÌœ »›« Ê—…
                    RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
           
                Else '€Ì— „ﬁÌœ »›« Ê—…
                    RSTransDetails("CostPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))
       
                End If
 
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
                    RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value * -1
                    RSTransDetails("OpeningSalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) * -1))
            
                    RSTransDetails("OpeningRESalesQty").value = RSTransDetails("Quantity").value
                    RSTransDetails("OpeningRESalesValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
                    RSTransDetails("Price").value = val(IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice")) = ""), 0, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemCostPrice"))))) / RSTransDetails("QtyBySmalltUnit").value
          
                End If

                RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
                RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
                RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))
                RSTransDetails("NoCount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("NoCount")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("NoCount"))))
                RSTransDetails("Width").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Width")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Width"))))
                RSTransDetails("Height").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Height")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Height"))))
                RSTransDetails("Area").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Area")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Area"))))
        
            'Dim TotalBillDiscount As Double
            'Dim TotalDiscountPerLine As Double
                  If Me.XPCboDiscountType.ListIndex = 1 Then
               ' TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.Text <> "" Then
               '     TotalBillDiscount = IIf(XPTxtDiscountVal.Text = "", Null, (XPTxtDiscountVal.Text)) * val(LblTotalAll.Caption) / 100
               '
                Else
               '     TotalBillDiscount = 0
                End If
            End If

      '      TotalDiscountPerLine = ((RSTransDetails("SHOWprice") * RSTransDetails("SHOWQTY")) / IIf(LblTotalAll.Caption = 0, 1, LblTotalAll.Caption)) * (TotalBillDiscount)
               '  TotalDiscountPerLine = FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / LblTotalAll.Caption * (TotalBillDiscount)
      
         '   RSTransDetails("TotalDiscountPerLine") = Round(TotalDiscountPerLine, 20) '2
                
                             Dim OldQty As Double
           '  Dim OldCost As Double
           '   Dim NewQty As Double
           '    Dim NewCost As Double
               
'getItemCostData XPDtbBill.value, RSTransDetails("Item_ID").value, val(DCboStoreName.BoundText), val(Me.XPTxtBillID.Text), OldQty, OldCost, NewQty, NewCost
'      RSTransDetails("OldQty").value = NewQty
'       RSTransDetails("OldCost").value = NewCost
'
'      RSTransDetails("NewQty").value = RSTransDetails("Quantity").value + RSTransDetails("OldQty").value
'       RSTransDetails("NewCost").value = ((RSTransDetails("OldQty").value * RSTransDetails("OldCost").value) + (RSTransDetails("Quantity").value * RSTransDetails("Price").value)) / (RSTransDetails("Quantity").value + RSTransDetails("OldQty").value)
       


                RSTransDetails.update
            End If

        Next RowNum
      
        Dim LngDevID As Long
        Dim LngDevNO  As Integer
        Dim StrTempAccountCode As String
        Dim StrTempDes As String
        Dim SngTemp  As Variant
           Dim SngTemp1  As Variant
           
         LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        '    '«·ﬁÌœ «·√Ê·
        '    '«·ÿ—› «·„œÌ‰
        '„—œÊœ«  «·„»Ì⁄« 

        'Transaction_Type=19 «–‰ ’—›
        'Transaction_Type=20 «–‰ «÷«›…

        Dim Account_Code_dynamic  As String
        Dim i As Integer
        '  SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
     '   SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) - ItemsGoodsTotalsnew '+ val(TxtValueAdded.Text)
        Dim Percetage As Double
PercentgValueAddedAccount_Transec XPDtbBill.value, 9, 0, , Percetage
   SngTemp = Round(SngTemp, 2)
        If SystemOptions.PriceWithVAT = True Then
            'SngTemp = SngTemp / 1.05
            SngTemp = SngTemp / (1 + (Percetage / 100))
        End If
        If SngTemp > 0 Then

                If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then
                    Account_Code_dynamic = get_account_code_branch(3, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» „—œÊœ«  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

              ' StrTempAccountCode = Account_Code_dynamic '„—œÊœ«  «·„»Ì⁄« 
            
            
            
            
         If val(DCDocTypes.BoundText) > 0 Then '„—œÊœ«  «·„»Ì⁄« 
                getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount

                        If StrTempAccountCode = "" And usedaccount = 1 Then
                                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·„—œÊœ«  «·„»Ì⁄«  ", vbCritical
                                    GoTo ErrTrap
                        ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                        
                        ElseIf usedaccount = 0 Then
                                StrTempAccountCode = Account_Code_dynamic '
                        End If

            Else
                        StrTempAccountCode = Account_Code_dynamic '
          End If
            
            
                StrTempDes = "„—œÊœ«  ⁄„·Ì…    —ﬁ„ " & Me.TxtNoteSerial1.Text
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
     '  If val(Me.TxtValueAdded.Text) <> 0 Then
     '      If SystemOptions.UserInterface = ArabicInterface Then
     '           StrTempDes = "«·ﬁÌ„… «·„÷«›… "
     '           Else
     '           StrTempDes = "VAT"
     '      End If
 '               LngDevNO = LngDevNO + 1
   'Dim AccountVATCreit As String
 'GetValueAddedAccount XPDtbBill.Value, , AccountVATCreit, 1, 9
 '               If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(Me.TxtValueAdded.Text), 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
 '                   GoTo ErrTrap
 '               End If
 '      End If
            ElseIf detect_inventory_work_type = 3 Then
                Dim groupAccount As String
             
                Dim line_value As Single

                With FG

                    For i = 1 To FG.Rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            ' groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 3)
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 3)

                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«» „—œÊœ«  «·„»Ì⁄«  ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
    
                            StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial1.Text
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If
 
      SngTemp1 = NewGrid.GetItemsTotal(ItemsServiceType) - ItemsServiceTotalsnew '+ val(TxtValueAdded.Text)
      SngTemp1 = Round(SngTemp1, 2)
If SystemOptions.PriceWithVAT = True Then
'SngTemp1 = SngTemp1 / 1.05
SngTemp1 = SngTemp1 / (1 + (Percetage / 100))


End If

        If SngTemp1 > 0 Then
        
                Account_Code_dynamic = get_account_code_branch(23, my_branch)
        
                If Account_Code_dynamic = "" Then
                          Account_Code_dynamic = get_account_code_branch(3, my_branch)
                End If
                
                StrTempAccountCode = Account_Code_dynamic '„—œÊœ«  «·„»Ì⁄« 
            
                StrTempDes = "„—œÊœ« «  ⁄„·Ì… —ﬁ„  —ﬁ„ " & Me.TxtNoteSerial1.Text
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp1, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
                

       
  End If
'       ''////////////////////
If SystemOptions.PriceWithVAT = True Then


'TxtValueAdded = (SngTemp + SngTemp1) * 0.05
TxtValueAdded = (SngTemp + SngTemp1) * Percetage / 100

End If
       If val(Me.TxtValueAdded.Text) <> 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "«·ﬁÌ„… «·„÷«›… "
                Else
                StrTempDes = "VAT"
           End If
                LngDevNO = LngDevNO + 1
  Dim AccountVATCreit As String
 GetValueAddedAccount XPDtbBill.value, AccountVATCreit, , 1, 9
                If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(Me.TxtValueAdded.Text), 0, StrTempDes & " „—œÊœ«  „»Ì⁄«   »—ﬁ„ " & TxtNoteSerial1 & CHR(13) & "  ·›« Ê—… «·»Ì⁄ —ﬁ„ " & "  " & TxtInvSerial, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
       End If
       '''///////////////
       SngTemp = SngTemp + val(Me.TxtValueAdded.Text)
        If (SngTemp + SngTemp1) > 0 Then
            If CboPayMentType.ListIndex = 0 Then
                '«·Œ“Ì‰…
                StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
                
                       If val(DCDocTypes.BoundText) > 0 Then '„—œÊœ«  «·„»Ì⁄« 
                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

                        If StrTempAccountCode = "" And usedaccount = 1 Then
                                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·„—œÊœ«  «·„»Ì⁄«  ", vbCritical
                                    GoTo ErrTrap
                        ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                        
                        ElseIf usedaccount = 0 Then
                                StrTempAccountCode = Account_Code_dynamic '
                        End If

            Else
                        StrTempAccountCode = StrTempAccountCode '
          End If
            
 
 
                StrTempDes = "„—œÊœ«  „»Ì⁄«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                'SngTemp = (Val(Me.XPTxtValue(0).text))
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Round(SngTemp + SngTemp1, 2), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
    
            If CboPayMentType.ListIndex = 1 Then
                '«·√Ã·
                StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
                                       If val(DCDocTypes.BoundText) > 0 Then '„—œÊœ«  «·„»Ì⁄« 
                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

                        If StrTempAccountCode = "" And usedaccount = 1 Then
                                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·„—œÊœ«  «·„»Ì⁄«  ", vbCritical
                                    GoTo ErrTrap
                        ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                        
                        ElseIf usedaccount = 0 Then
                                StrTempAccountCode = Account_Code_dynamic '
                        End If

            Else
                        StrTempAccountCode = StrTempAccountCode '
          End If


                StrTempDes = "„—œÊœ«  „»Ì⁄«  —ﬁ„ " & Me.TxtNoteSerial1.Text
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Round(SngTemp + SngTemp1, 2), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
    
        End If
    'noteid1 = val(general_noteid)
                updateNotesValueAndNobytext CDbl(general_noteid), CDbl(Round(SngTemp + SngTemp1, 2))
                 
                 
        ' If CboRetrunType.ListIndex = 0 Then
        'create_recieve_voucher
        ' End If
 
        'If SystemOptions.autoReseiveVoucher = True Then

        'End If
        If SystemOptions.USERautoIssueVoucher = False Then
        
                     If SystemOptions.returnnotcreatvoucher = False Then
                               If Not CreateRecieveVoucher Then BeginTrans = True: MsgBox "ÕœÀ Œÿ√ «À‰«¡ «‰‘«¡ «–‰ «·«” ·«„ ": GoTo ErrTrap
                      End If
        
        End If
'Wael

'
'SngTemp = Round(SngTemp, 2)
'
'If SystemOptions.PriceWithVAT = True Then
''SngTemp = SngTemp / 1.05
''WaelNew
'If chkTaxExempt.value = vbChecked Then
'        SngTemp = SngTemp
'    Else
'        SngTemp = SngTemp / (1 + (Percetage / 100))
'    End If
'End If
'
'
'
'End If
'        If SngTemp > 0 Then
'
'            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then
'                Account_Code_dynamic = get_account_code_branch(3, my_branch)
'
'                If Account_Code_dynamic = "NO branch" Then
'                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'                    GoTo ErrTrap
'                Else
'
'                    If Account_Code_dynamic = "NO account" Then
'                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» „—œÊœ«  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                        GoTo ErrTrap
'
'                    End If
'                End If
'
'              ' StrTempAccountCode = Account_Code_dynamic '„—œÊœ«  «·„»Ì⁄« 
'
'
'
'
'         If val(DCDocTypes.BoundText) > 0 Then '„—œÊœ«  «·„»Ì⁄« 
'                getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount
'
'                        If StrTempAccountCode = "" And usedaccount = 1 Then
'                                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·„—œÊœ«  «·„»Ì⁄«  ", vbCritical
'                                    GoTo ErrTrap
'                        ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
'
'                        ElseIf usedaccount = 0 Then
'                                StrTempAccountCode = Account_Code_dynamic '
'                        End If
'
'            Else
'                        StrTempAccountCode = Account_Code_dynamic '
'          End If
'
'
'                StrTempDes = "„—œÊœ«  ⁄„·Ì…    —ﬁ„ " & Me.TxtNoteSerial1.Text
'                LngDevNO = LngDevNO + 1
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'     '  If val(Me.TxtValueAdded.Text) <> 0 Then
'     '      If SystemOptions.UserInterface = ArabicInterface Then
'     '           StrTempDes = "«·ﬁÌ„… «·„÷«›… "
'     '           Else
'     '           StrTempDes = "VAT"
'     '      End If
' '               LngDevNO = LngDevNO + 1
'   'Dim AccountVATCreit As String
' 'GetValueAddedAccount XPDtbBill.Value, , AccountVATCreit, 1, 9
' '               If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(Me.TxtValueAdded.Text), 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.Value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
' '                   GoTo ErrTrap
' '               End If
' '      End If
'            ElseIf detect_inventory_work_type = 3 Then
'                Dim groupAccount As String
'
'                Dim line_value As Single
'
'                With FG
'
'                    For i = 1 To FG.Rows - 1
'
'                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
'
'                            ' groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 3)
'                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 3)
'
'                            If groupAccount = "Error" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«» „—œÊœ«  «·„»Ì⁄«  ·„Ã„Ê⁄ …"
'                                Else
'                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
'                                End If
'
'                                GoTo ErrTrap
'                            End If
'
'                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count"))
'
'                            StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial1.Text
'                            LngDevNO = LngDevNO + 1
'
'                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                                GoTo ErrTrap
'                            End If
'
'                        End If
'
'                    Next i
'
'                End With
'
'            End If
'        End If
'
'      SngTemp1 = NewGrid.GetItemsTotal(ItemsServiceType) - ItemsServiceTotalsnew '+ val(TxtValueAdded.Text)
'      SngTemp1 = Round(SngTemp1, 2)
'
'
'If SystemOptions.PriceWithVAT = True Then
''SngTemp1 = SngTemp1 / 1.05
''WaelNew
'        If chkTaxExempt.value = vbChecked Then
'            SngTemp1 = SngTemp1
'        Else
'            SngTemp1 = SngTemp1 / (1 + (Percetage / 100))
'        End If
'
'End If
'        If SngTemp1 > 0 Then
'
'                Account_Code_dynamic = get_account_code_branch(23, my_branch)
'
'                If Account_Code_dynamic = "" Then
'                          Account_Code_dynamic = get_account_code_branch(3, my_branch)
'                End If
'
'                StrTempAccountCode = Account_Code_dynamic '„—œÊœ«  «·„»Ì⁄« 
'
'                StrTempDes = "„—œÊœ« «  ⁄„·Ì… —ﬁ„  —ﬁ„ " & Me.TxtNoteSerial1.Text
'                LngDevNO = LngDevNO + 1
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp1, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'
'
'
'        End If
''       ''////////////////////
'
'If SystemOptions.PriceWithVAT = True Then
'
'
''TxtValueAdded = (SngTemp + SngTemp1) * 0.05
''WaelNew
'    If chkTaxExempt.value = vbChecked Then
'        TxtValueAdded = 0
'    Else
'        TxtValueAdded = (SngTemp + SngTemp1) * Percetage / 100
'    End If
'
'End If
'        Dim AccountVATCreit As String
'       If val(Me.TxtValueAdded.Text) <> 0 Then
'           If SystemOptions.UserInterface = ArabicInterface Then
'                StrTempDes = "«·ﬁÌ„… «·„÷«›… "
'                Else
'                StrTempDes = "VAT"
'           End If
'                LngDevNO = LngDevNO + 1
'
'                 GetValueAddedAccount XPDtbBill.value, AccountVATCreit, , 1, 9
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(Me.TxtValueAdded.Text), 0, StrTempDes & " „—œÊœ«  „»Ì⁄«   »—ﬁ„ " & TxtNoteSerial1 & CHR(13) & "  ·›« Ê—… «·»Ì⁄ —ﬁ„ " & "  " & TxtInvSerial, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'       End If
'
'       '''///////////////
'       SngTemp = SngTemp + val(Me.TxtValueAdded.Text)
'        If (SngTemp + SngTemp1) > 0 Then
'            If CboPayMentType.ListIndex = 0 Then
'                '«·Œ“Ì‰…
'                    StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
'
'                       If val(DCDocTypes.BoundText) > 0 Then '„—œÊœ«  «·„»Ì⁄« 
'                                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount
'
'                            If StrTempAccountCode = "" And usedaccount = 1 Then
'                                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·„—œÊœ«  «·„»Ì⁄«  ", vbCritical
'                                        GoTo ErrTrap
'                            ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
'
'                            ElseIf usedaccount = 0 Then
'                                    StrTempAccountCode = Account_Code_dynamic '
'                            End If
'
'                            Else
'                                      StrTempAccountCode = StrTempAccountCode '
'                            End If
'                   End If
'
'
'                StrTempDes = "„—œÊœ«  „»Ì⁄«  —ﬁ„ " & Me.TxtNoteSerial1.Text
'                'SngTemp = (Val(Me.XPTxtValue(0).text))
'                LngDevNO = LngDevNO + 1
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Round(SngTemp + SngTemp1, 2), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'      '  End If
'            If CboPayMentType.ListIndex = 1 Then
'                '«·√Ã·
'                StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
'                                       If val(DCDocTypes.BoundText) > 0 Then '„—œÊœ«  «·„»Ì⁄« 
'                getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount
'
'                        If StrTempAccountCode = "" And usedaccount = 1 Then
'                                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·„—œÊœ«  «·„»Ì⁄«  ", vbCritical
'                                    GoTo ErrTrap
'                        ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
'
'                        ElseIf usedaccount = 0 Then
'                                StrTempAccountCode = Account_Code_dynamic '
'                        End If
'
'            Else
'                        StrTempAccountCode = StrTempAccountCode '
'          End If
'
'
'                StrTempDes = "„—œÊœ«  „»Ì⁄«  —ﬁ„ " & Me.TxtNoteSerial1.Text
'                LngDevNO = LngDevNO + 1
'
'                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Round(SngTemp + SngTemp1, 2), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.Text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'        End If
'    'noteid1 = val(general_noteid)
'                updateNotesValueAndNobytext CDbl(general_noteid), CDbl(Round(SngTemp + SngTemp1, 2))
'
'
'        ' If CboRetrunType.ListIndex = 0 Then
'        'create_recieve_voucher
'        ' End If
'
'        'If SystemOptions.autoReseiveVoucher = True Then
'
'        'End If
'        If SystemOptions.USERautoIssueVoucher = False Then
'
'                     If SystemOptions.returnnotcreatvoucher = False Then
'                               If Not CreateRecieveVoucher Then BeginTrans = True: MsgBox "ÕœÀ Œÿ√ «À‰«¡ «‰‘«¡ «–‰ «·«” ·«„ ": GoTo ErrTrap
'                      End If
'
'        End If
'

   SaveItemsData
   SaveValueAdded
   

'************************************************************************************
   Set RSTransDetails1 = New ADODB.Recordset
   StrSQL = "SELECT   * from dbo.TblTransactionPayments Where (1 = -1)"
   RSTransDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
 
                                    'Check Repeat Serial
                                         
                            If val(Me.DcboBox.BoundText) <> 0 And val(txtPPointID) <> 0 Then
                                           RSTransDetails1.AddNew
                                            RSTransDetails1("boxid").value = val(Me.DcboBox.BoundText)
                                            RSTransDetails1("Recorddate").value = XPDtbBill.value
                                           RSTransDetails1("PointID").value = val(txtPPointID)
                                           RSTransDetails1("CurrentCashireID").value = CurrentCashireID
                                           
                                           RSTransDetails1("Transaction_ID").value = val(XPTxtBillID.Text)
                                           RSTransDetails1("PaymentID").value = 0
                                          If val(txtPPointID) <> 0 Then
                                           RSTransDetails1("Value").value = val(LblTotal.Caption)
                                           End If
                                     '     RSTransDetails1("Value").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("Value")) = ""), 0, val(Grid.TextMatrix(RowNum, Grid.ColIndex("Value"))))
                                          ' RSTransDetails1("CardNo").value = IIf((Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo")) = ""), "", (Grid.TextMatrix(RowNum, Grid.ColIndex("CardNo"))))
                                           
                                             
                                           RSTransDetails1("CardNo").value = ""
                                            RSTransDetails1("Effect").value = -1
                                           RSTransDetails1.update
                                  
 End If
'***************************************************************************************
saveBillBuy

        Cn.CommitTrans
        BeginTrans = False

        'salimher 09042019
    
        '----------------------------------------------------------------
        '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
        StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=9" '& InvType
         
        If SystemOptions.usertype <> UserAdminAll Then
            StrSQL = StrSQL & " AND   BranchId=" & Current_branch
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Me.Retrive val(Me.XPTxtBillID.Text)
        '----------------------------------------------------------------
        CuurentLogdata
 
        Select Case Me.TxtModFlg.Text

            Case "N"
        
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & CHR(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            
                If SystemOptions.Save_options = 1 Or SystemOptions.Save_options = 2 Then
                    PrintReport 0

                    DoEvents
                    DoEvents
                    DoEvents
        
                ElseIf SystemOptions.Save_options = 3 Then
                    PrintReport 0

                    DoEvents
                    DoEvents
                    DoEvents
        
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
        
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                lbl(11).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
            
               
        End Select

        TxtModFlg.Text = "R"
        Me.Retrive val(Me.XPTxtBillID.Text)
    End If

    Screen.MousePointer = vbDefault
    ' '----------------------------------------------------------------
    ' '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
    ' StrSQL = "SELECT * FROM Transactions WHERE Transaction_Type=9"
    '
    ' Set rs = New ADODB.Recordset
    ' rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    ' Me.Retrive Val(Me.XPTxtBillID.text)
    '----------------------------------------------------------------
    
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub
Sub BillCustomer()
Dim Msg As String
If Me.TxtModFlg.Text <> "R" Then
If Me.TxtModFlg.Text = "N" Then
RetriveBillBuy val(DBCboClientName.BoundText)
End If
If Me.TxtModFlg.Text = "E" And (FlgBillBuy = True Or VSFlexGrid1.Rows = 1) Then
RetriveBillBuy val(DBCboClientName.BoundText)
End If
End If
End Sub
Sub RetriveBillBuy(Optional CuID As Double = 0)
Dim sql As String
Dim Rs8 As ADODB.Recordset
Dim i As Integer
Set Rs8 = New ADODB.Recordset
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.Rows = 1
End With
If 1 = 1 Then
sql = " SELECT      dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
sql = sql & "                      dbo.Transactions.ManualNO, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.CusID,"
sql = sql & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.TotalPayed, dbo.Transactions.OldContID,"
sql = sql & "                      dbo.transactions.OldValue , dbo.transactions.dueDate, dbo.transactions.Vat, dbo.transactions.Transaction_NetValue"
sql = sql & " FROM         dbo.Transactions LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
sql = sql & "  WHERE     (dbo.Transactions.PaymentType = 1) AND (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                       dbo.Transactions.Transaction_Type = 2 or dbo.Transactions.Transaction_Type = 71) AND (dbo.Transactions.TotalPayed IS NULL OR"
sql = sql & "                       dbo.Transactions.TotalPayed = 0) AND (dbo.Transactions.CusID = " & CuID & ")"

If val(CboRetrunType.ListIndex) = 0 Then
sql = sql & " AND (dbo.Transactions.NoteSerial1 = '" & TxtInvSerial.Text & "')"
End If
sql = sql & "  ORDER BY dbo.Transactions.DueDate ,dbo.Transactions.NoteSerial1"

Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
VSFlexGrid1.Enabled = True


        VSFlexGrid1.Enabled = True
With VSFlexGrid1
.Clear flexClearScrollable, flexClearEverything
.Rows = 1
    .Rows = .Rows + Rs8.RecordCount
.Rows = .FixedRows + Rs8.RecordCount
Rs8.MoveFirst
For i = .FixedRows To Rs8.RecordCount
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(Rs8("BranchId").value), 0, Rs8("BranchId").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_name").value), 0, Rs8("branch_name").value)
Else
.TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(Rs8("branch_namee").value), 0, Rs8("branch_namee").value)
End If

.TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(Rs8("DueDate").value), "", Rs8("DueDate").value)
.TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(Rs8("Transaction_ID").value), 0, Rs8("Transaction_ID").value)
.TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(Rs8("Transaction_Date").value), "", Rs8("Transaction_Date").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs8("NoteSerial1").value), "", Rs8("NoteSerial1").value)
.TextMatrix(i, .ColIndex("too")) = IIf(IsNull(Rs8("ManualNO").value), "", Rs8("ManualNO").value)
.TextMatrix(i, .ColIndex("Note_Value")) = val(IIf(IsNull(Rs8("Transaction_NetValue").value), IIf(IsNull(Rs8("OldValue").value), 0, Rs8("OldValue").value), Rs8("Transaction_NetValue").value))
If val(.TextMatrix(i, .ColIndex("NoteID"))) <> 0 Then
.TextMatrix(i, .ColIndex("PayedValue")) = GeteBillBuy(val(.TextMatrix(i, .ColIndex("NoteID"))))
Else
.TextMatrix(i, .ColIndex("PayedValue")) = 0
End If
.TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("Note_Value"))) - val(.TextMatrix(i, .ColIndex("PayedValue")))
Rs8.MoveNext
Next i
End With
End If
End If
End Sub
Public Sub RetriveBillBuyData(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrSQL As String


   ' On Error GoTo ErrTrap
    Set RsDetails = New ADODB.Recordset
  StrSQL = "   SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblNotesBillBuyPayment2.*"
  StrSQL = StrSQL & "  FROM         dbo.TblNotesBillBuyPayment2 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblNotesBillBuyPayment2.branch_no = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & "  Where (dbo.TblNotesBillBuyPayment2.NoteID1 = " & val(XPTxtBillID.Text) & " and dbo.TblNotesBillBuyPayment2.TransType=1)"
    
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With VSFlexGrid1
    .Clear flexClearScrollable, flexClearEverything
    .Rows = .FixedRows

    If Not (RsDetails.BOF Or RsDetails.EOF) Then
        RsDetails.MoveFirst
        .Rows = .FixedRows + RsDetails.RecordCount

        For i = .FixedRows To RsDetails.RecordCount
        .TextMatrix(i, .ColIndex("Ser")) = i

            .TextMatrix(i, .ColIndex("branch_no")) = IIf(IsNull(RsDetails("branch_no").value), 0, RsDetails("branch_no").value)
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_name").value), "", RsDetails("branch_name").value)
            Else
            .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(RsDetails("branch_namee").value), 0, RsDetails("branch_namee").value)
            End If
            .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsDetails("NoteID").value), 0, RsDetails("NoteID").value)
            .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDetails("NoteSerial1").value), 0, RsDetails("NoteSerial1").value)
            .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsDetails("Note_Value").value), 0, RsDetails("Note_Value").value)
            .TextMatrix(i, .ColIndex("PayedValue")) = IIf(IsNull(RsDetails("PayedValue").value), 0, RsDetails("PayedValue").value)
            .TextMatrix(i, .ColIndex("TransPayedValue")) = IIf(IsNull(RsDetails("TransPayedValue").value), 0, RsDetails("TransPayedValue").value)
            .TextMatrix(i, .ColIndex("too")) = IIf(IsNull(RsDetails("too").value), "", RsDetails("too").value)
            .TextMatrix(i, .ColIndex("NetValue")) = IIf(IsNull(RsDetails("NetValue").value), 0, RsDetails("NetValue").value)
            .TextMatrix(i, .ColIndex("RemainingValue")) = IIf(IsNull(RsDetails("RemainingValue").value), 0, RsDetails("RemainingValue").value)
            .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(((RsDetails("DueDate").value))), " ", ((RsDetails("DueDate").value)))
            .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(((RsDetails("NoteDate").value))), "", ((RsDetails("NoteDate").value)))
            .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked
            RsDetails.MoveNext
        Next i
        

    End If
End With
RelineBuy
    RsDetails.Close
    Set RsDetails = Nothing
    Set rs = Nothing
    Exit Sub
ErrTrap:
End Sub
Function GeteBillBuy(Optional Transaction_ID As Double = 0) As Double
Dim sql As String
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
sql = " SELECT   SUM(PayedValue) AS Smatiobn"
sql = sql & " From dbo.TblBillBuyPayment2"
sql = sql & " Where (Transaction_ID = " & Transaction_ID & ")"
sql = sql & " GROUP BY Transaction_ID"
Rs8.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
GeteBillBuy = IIf(IsNull(Rs8("Smatiobn").value), 0, Rs8("Smatiobn").value)
Else
GeteBillBuy = 0
End If
End Function
Function CheckFilegrid() As Boolean
If CboRetrunType.ListIndex = 1 Then CheckFilegrid = True: Exit Function
Dim i As Integer
Dim j As Integer
Dim Item_ID As Double
Dim SumQty As Double
Dim ClassId As Integer
Dim itemsize As Integer
Dim ColorID As Integer
Dim UnitID As Integer
Dim total As Double
Dim Msg As String
With FG
CheckFilegrid = True
For j = .FixedRows To .Rows - 1

SumQty = 0
Item_ID = val(.TextMatrix(j, .ColIndex("Code")))
ClassId = val(.TextMatrix(j, .ColIndex("ClassId")))
itemsize = val(.TextMatrix(j, .ColIndex("ItemSize")))
ColorID = val(.TextMatrix(j, .ColIndex("ColorID")))
UnitID = IIf(.Cell(flexcpData, j, .ColIndex("UnitID")) = "", 0, (.Cell(flexcpData, j, .ColIndex("UnitID"))))
For i = .FixedRows To .Rows - 1

If Item_ID = val(.TextMatrix(i, .ColIndex("Code"))) And UnitID = IIf(.Cell(flexcpData, i, .ColIndex("UnitID")) = "", 0, (.Cell(flexcpData, i, .ColIndex("UnitID")))) And ClassId = val(.TextMatrix(i, .ColIndex("ClassId"))) And itemsize = val(.TextMatrix(i, .ColIndex("ItemSize"))) And ColorID = val(.TextMatrix(i, .ColIndex("ColorID"))) Then
SumQty = SumQty + val(.TextMatrix(i, .ColIndex("Count")))
End If
Next i
total = RetriveQtyItem(TxtInvSerial.Text, Item_ID, ColorID, ClassId, itemsize, UnitID)
If total < SumQty Then
If total > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
Msg = .Cell(flexcpTextDisplay, j, .ColIndex("Name")) & "  ·«Ì„ﬂ‰ «— Ã«⁄ ﬂ„Ì… «ﬂ»— „‰ «·ﬂ„Ì… «·«’·Ì… ··’‰› "
Msg = Msg & CHR(13)
Msg = Msg & (total) & " " & "«·ﬂ„Ì… «·„ »ﬁÌ…"
Else
End If
Else
If SystemOptions.UserInterface = ArabicInterface Then
Msg = .Cell(flexcpTextDisplay, j, .ColIndex("Name")) & "  ·«ÌÊÃœ  ﬂ„Ì… „‰  «·’‰›  "
Msg = Msg & CHR(13)
Msg = Msg & "·«— Ã«⁄Â«"
Else
End If
End If
MsgBox Msg
GoTo l
Else
CheckFilegrid = True
End If
Next j
Exit Function
End With
l: CheckFilegrid = False


End Function
Function create_recieve_voucher()
    Dim Transaction_serial As Integer
    Dim MYTEXT As String
    Transaction_serial = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
    MYTEXT = TxtTransSerial

    Dim Transaction_ID As Integer
    Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))

    Cn.Execute "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,ManualNO)SELECT " & Transaction_ID & "," & MYTEXT & ",Transaction_Date,Transaction_Type = 20,CusID,StoreID,UserID,Emp_ID,nots=1,ManualNO From Transactions Where Transaction_ID =" & XPTxtBillID.Text + " And Transaction_Type = 9"
    '
    Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,UnitId,ShowQty,QtyBySmalltUnit)SELECT round(showPrice + ToTAlELSHahn/ShowQty,2),guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, costprice, ColorID, UnitId, ShowQty, QtyBySmalltUnit From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.Text

    'CreateglForVoucher
 
End Function
 
Private Sub XPCboDiscountType_Change()
XPCboDiscountType_Click
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
    
        XPTxtDiscountVal.Enabled = False
        XPTxtDiscountVal.Text = ""
    Else
    
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.Text = ""
    End If

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1, , , True
        End If
    End If

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    'Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)
    If XPCboDiscountType.ListIndex = 0 Then
        'lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
    '    lbl(8).Visible = False
    Else
        'lbl(8).Visible = True
        XPTxtDiscountVal.Visible = True
    '    lbl(8).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.Text = "N" Then
                    XPTxtValue(0).Text = ""
                    XPTxtSerial(0).Text = ""
                End If

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).Text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.Text = "N" Then
                    XPTxtValue(1).Text = ""
                    XPTxtSerial(1).Text = ""
                    DtpDelayDate.value = Date
                    XPTxtSerial(1).Text = CStr(new_id("Notes", "NoteSerial", "", True))
                End If

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                
                End If

            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).Text = ""
                '            XPTxtSerial(1).Enabled = False
                XPTxtValue(1).Text = ""
            End If

        Case 2

            If XPChkPayType(2).value = Checked Then
                If Me.TxtModFlg.Text = "N" Then
                    XPTxtValue(2).Text = ""
                    XPTxtChqueNum.Text = ""
                    XPDTPDueDate.value = Date
                    DCboBankName.BoundText = ""
                End If

                If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
                    XPTxtValue(2).Enabled = True
                    XPTxtChqueNum.Enabled = True
                    XPDTPDueDate.Enabled = True
                    XPTxtValue(2).locked = False
                    XPTxtChqueNum.locked = False
                    DCboBankName.locked = False
                    DCboBankName.Enabled = True
                End If

            Else
                XPTxtValue(2).Text = ""
                XPTxtValue(2).Enabled = False
                XPTxtChqueNum.Enabled = False
                XPDTPDueDate.Enabled = False
                DCboBankName.locked = True
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub PrintReport(Optional repType As Integer)
    On Error GoTo ErrTrap
    Dim ShowType As Integer
    Dim SaleReport As ClsRepoerts
    Dim StrPath As String
    Dim Msg As String
    Dim Fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    
    If XPTxtBillID.Text <> "" Then
        Set SaleReport = New ClsRepoerts
        
Dim X As Integer
 
   X = MsgBox("ÿ»«⁄Â „»«‘—Â", vbInformation + vbYesNo)
   
 
    
        
        SaleReport.ReturnSallingData XPTxtBillID.Text, Round(val(LblTotal), 2), Fullcode, , val(Me.dcBranch.BoundText), repType, X
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.Text) <> "" Then
        oldtxtNoteSerial1.Text = TxtNoteSerial1.Text
    End If

    TxtNoteSerial.Text = ""
    TxtNoteSerial1.Text = ""
    CurrentVoucherNo = ""
TxtNoteSerial1V = ""
    DateChanged = True

End Sub

Private Sub XPTxtDiscountVal_Change()
    Dim Msg As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        NewGrid.Calculate 1, , , True
    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.Text, 0)
End Sub

Private Sub XPTxtSum_Change()

    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).Text = XPTxtSum.Text
    End If
RelinVatGrid
   ' Me.LblTotal.Caption = XPTxtSum.Text
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    If Trim(Me.TxtModFlg.Text) = "" Then Exit Sub

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
ErrTrap:         End Sub

Public Sub Convert()
    Cmd_Click (0)
End Sub

Public Sub Cala()
    NewGrid.Calculate 1, , , True
End Sub

Private Sub DBCboClientName_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    Dim Fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.Text = Fullcode
    
    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                CboPayMentType.locked = True
                CboPayMentType.ListIndex = 0
            Else
                CboPayMentType.locked = False
            End If
                    
        End If
                
        Dim DefaultSalesPersonId As Integer
        GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId

        If Not DefaultSalesPersonId = 0 Then

            Me.DcboEmp.BoundText = DefaultSalesPersonId
        End If
            StrSQL = "Select * From TblCustemers Where CusID=" & val(DBCboClientName.BoundText)
            rs2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If rs2.RecordCount > 0 Then
            Me.TxtVATNO.Text = IIf(IsNull(rs2("VATNO").value), "", rs2("VATNO").value)
            Else
            Me.TxtVATNO.Text = ""
            End If

    End If
If SystemOptions.AllowReturnFIFO = True Then
BillCustomer
End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Function CheckInvData() As Boolean
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String

    CheckInvData = True
    Exit Function

    If Me.TxtInvSerial.Text <> "" Then
        StrSQL = "SELECT * From Transactions "
        StrSQL = StrSQL + " Where Transaction_Serial='" & Trim(Me.TxtInvSerial.Text) & "'"
        StrSQL = StrSQL + " AND (Transactions.Transaction_Type=2 or Transactions.Transaction_Type=21) "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If rs.BOF Or rs.EOF Then
            Msg = "·« ÊÃœ ›« Ê—… »Ì⁄ »Â–« «·—ﬁ„..!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CheckInvData = False
            rs.Close
            Set rs = Nothing
            Exit Function
        ElseIf rs("CusID").value <> Me.DBCboClientName.BoundText Then
            Msg = "«·›« Ê—… —ﬁ„ " & Trim(Me.TxtInvSerial.Text)
            Msg = Msg & CHR(13) & "·Ì”  „⁄ «·⁄„Ì·" & Me.DBCboClientName.Text
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            CheckInvData = False
            rs.Close
            Set rs = Nothing
            Exit Function
        Else
        
            Me.TxtInvID.Text = rs("Transaction_ID").value
        End If
    End If

    rs.Close
    Set rs = Nothing
    CheckInvData = True
End Function

Private Function CheckRetrunInv() As Boolean
    Dim StrSQL  As String
    Dim rs As New ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    '----------------------------
If CboRetrunType.ListIndex = 1 Then
CheckRetrunInv = True
Exit Function
End If
    StrSQL = "Select * From Transaction_Details Where Transaction_ID=" & val(Me.TxtInvID.Text) & ""
    StrSQL = StrSQL + " Order  By ID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    CheckRetrunInv = False

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FG

            For i = .FixedRows To .Rows - 1

                If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                    If rs.filter <> adFilterNone Then
                        rs.filter = adFilterNone
                    End If

                    rs.MoveFirst
                    rs.filter = "Item_ID=" & val(.TextMatrix(i, .ColIndex("Name")))

                    If rs.BOF Or rs.EOF Then
                        Msg = "«·’‰› : " & .Cell(flexcpTextDisplay, i, .ColIndex("Name"))
                        Msg = Msg & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„ : " & i
                        Msg = Msg & CHR(13) & "·„ Ìﬂ‰ „ÊÃÊœ ›Ï «·›« Ê—… —ﬁ„ : " & Me.TxtInvSerial.Text
                        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        CheckRetrunInv = False
                        rs.Close
                        Set rs = Nothing
                        Exit Function
                    ElseIf FG.Cell(flexcpChecked, i, .ColIndex("HaveSerial")) = flexChecked Then
                        rs.find "ItemSerial='" & Trim(.TextMatrix(i, .ColIndex("Serial"))) & "'", , adSearchForward, 1

                        If rs.BOF Or rs.EOF Then
                            Msg = "«·ﬁÿ⁄… –«  «·”Ì—Ì«·:  " & Trim(.TextMatrix(i, .ColIndex("Serial")))
                            Msg = Msg & CHR(13) & "„‰ «·’‰› : " & .Cell(flexcpTextDisplay, i, .ColIndex("Name"))
                            Msg = Msg & CHR(13) & "Ê«·„ÊÃÊœ ›Ï «·”ÿ— —ﬁ„  : " & i
                            Msg = Msg & CHR(13) & "·„ Ìﬂ‰ „ÊÃÊœ ›Ï «·›« Ê—… —ﬁ„  : " & Me.TxtInvSerial.Text
                            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                            CheckRetrunInv = False
                            rs.Close
                            Set rs = Nothing
                            Exit Function
                        End If
                    End If
                End If

            Next i

        End With

    End If

    '----------------------------

    '----------------------------
    CheckRetrunInv = True
End Function
